import { async } from "regenerator-runtime";

//import { get } from "core-js/core/dict"
const axios = require("axios");

const OpenAI = require("openai");

const openai = new OpenAI({
  apiKey: "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",
  dangerouslyAllowBrowser: true,
  version: "v1",
});

const submitButton = document.getElementById("submit");
const output = document.getElementById("output");
const inputElement = document.querySelector("input");
const history = document.querySelector(".history");
const buttonElement = document.querySelector("button");
const conversationHistory = [{ role: "system", content: "You are a helpful assistant." }];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
      if (result.status === "succeeded") {
        const accessToken = result.value;

        const assistant = await openai.beta.assistants.create({
          // See if you can come up with a better instruction
          instructions:
            "You are an Outlook QA assistant, who answers questions by making calls to Outlook to retrieve user emails and using that information to successfully answer the users questions. When the user types lets start fresh you clear all messages in the thread.",
          name: "Outlook QA Assistant",
          model: "gpt-3.5-turbo-1106",
          tools: [
            {
              type: "function",
              function: {
                name: "getAllMessages",
                description:
                  "Get the user's emails from Outlook based on a search keyword and/or sender email address and/or number of emails to retrieve",
                parameters: {
                  type: "object",
                  properties: {
                    accessToken: {
                      type: "string",
                      description: "The access token for the user's Outlook account. Required for authentication.",
                    },
                    searchKeyword: {
                      type: "string",
                      description:
                        "The keyword to search for in the user's emails and return all emails that match the keyword.",
                    },
                    from: {
                      type: "string",
                      description:
                        "The sender email address to search for in the user's emails and return all emails that match the sender email address.",
                      unit: {
                        type: "string",
                        pattern: "^\\S+@\\S+\\.\\S+$",
                      },
                    },
                    top: {
                      type: "integer",
                      description: "The number of emails to retrieve from the user's emails.",
                    },
                  },
                  required: ["accessToken"],
                },
              },
            },
          ],
        });

        const thread = await openai.beta.threads.create();

        submitButton.addEventListener("click", () => getMessage(assistant.id, accessToken, thread.id));

        buttonElement.addEventListener("click", clearInput);
      } else {
        console.error(result.error);
      }
    });
  }
});

function changeInput(value) {
  inputElement.value = value;
}

function clearInput() {
  inputElement.value = "";
}

// function maskingEmailBody(textContent) {
//   return new Promise((resolve, reject) => {
//     try {
//       // Use string interpolation to include textContent in the command
//       const command = `python masking.py "${textContent}"`;

//       exec(command, (error, stdout, stderr) => {
//         if (error) {
//           console.error("Error executing Python script:", error.message);
//           reject(error);
//         } else {
//           console.log("Python script output:", stdout);
//           // Directly return the masked text
//           resolve(stdout.trim());
//         }
//       });
//     } catch (error) {
//       console.error("Exception:", error.message);
//       reject(error);
//     }
//   });
// }

function getTextFromHtml(htmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, "text/html");
  const textContent = doc.body.textContent;
  return maskingEmailBody(textContent);
}

// async function getTextFromHtml(htmlString) {
//   return new Promise((resolve, reject) => {
//     try {
//       const parser = new DOMParser();
//       const doc = parser.parseFromString(htmlString, "text/html");
//       const textContent = doc.body.textContent;

//       // Call the asynchronous masking function with textContent as an argument
//       maskingEmailBody(textContent)
//         .then((maskedText) => {
//           resolve(maskedText);
//         })
//         .catch((error) => {
//           reject(error);
//         });
//     } catch (error) {
//       reject(error);
//     }
//   });
// }

function parseCCRecipients(ccRecipients) {
  let ccRecipientsArray = [];
  if (ccRecipients && ccRecipients.length > 0) {
    ccRecipients.forEach(function (ccRecipient) {
      let cc = {
        name: ccRecipient.EmailAddress.Name,
        emailAddress: ccRecipient.EmailAddress.Address,
      };
      ccRecipientsArray.push(cc);
    });
  }
  return ccRecipientsArray;
}

function getAllMessages(params) {
  let { accessToken, searchKeyword = "", from = "None", top = 500 } = params;
  return new Promise(async (resolve, reject) => {
    let messages = [];
    let url;

    if (searchKeyword === "" && from !== "None") {
      url =
        Office.context.mailbox.restUrl + `/v2.0/me/messages?$filter=from/emailAddress/address eq '${from}'&$top=${top}`;
    } else if (searchKeyword !== "" && from === "None") {
      url =
        Office.context.mailbox.restUrl + `/v2.0/me/messages?$search=${encodeURIComponent(searchKeyword)}&$top=${top}`;
    } else if (searchKeyword !== "" && from !== "None") {
      url =
        Office.context.mailbox.restUrl +
        `/v2.0/me/messages?$search=${encodeURIComponent(
          searchKeyword
        )}&$filter=from/emailAddress/address eq '${from}'&$top=${top}`;
    }

    async function fetchMessages(url) {
      try {
        const response = await axios.get(url, {
          headers: {
            Authorization: "Bearer " + accessToken,
            Prefer: 'outlook.body-content-type="text"',
            InferenceClassification: "Focused",
          },
        });

        //console.log(response.data);
        //console.log(response.data.value.length);

        if (response.data.value && response.data.value.length > 0) {
          response.data.value.forEach(function (message) {
            let messagejson = {
              subject: message.Subject,
              body: getTextFromHtml(message.Body.Content).replace(/\s+/g, " ").trim(),
              sender: message.Sender.EmailAddress.Address,
              receivedDateTime: message.ReceivedDateTime,
              id: message.Id,
              importance: message.Importance,
              isRead: message.IsRead,
              isReadReceiptRequested: message.IsReadReceiptRequested,
              CCRecipients: parseCCRecipients(message.CcRecipients),
              BCCRecipients: parseCCRecipients(message.BccRecipients),
            };
            messages.push(messagejson);
          });
        }

        if (response.data["@odata.nextLink"]) {
          await fetchMessages(response.data["@odata.nextLink"]);
        } else {
          resolve(messages); // Resolve the promise when all messages are fetched
        }
      } catch (error) {
        console.error("Error from Axios:", error);
        // Stop the call and return the specified string on error
        resolve(messages);
      }
    }

    fetchMessages(url);
  });
}

export async function getMessage(assistant_id, accessToken, thread_id) {
  console.log("Clicked");

  const message = await openai.beta.threads.messages.create(thread_id, {
    role: "user",
    content: inputElement.value,
  });

  const run = await openai.beta.threads.runs.create(thread_id, {
    assistant_id: assistant_id,
    //instructions: "Please address the user as Jane Doe. The user has a premium account."
  });

  let runstatus = await openai.beta.threads.runs.retrieve(thread_id, run.id);

  while (runstatus.status !== "completed") {
    await new Promise((resolve) => setTimeout(resolve, 1000));
    runstatus = await openai.beta.threads.runs.retrieve(thread_id, run.id);

    if (runstatus.status === "requires_action") {
      const toolCalls = runstatus.required_action.submit_tool_outputs.tool_calls;
      const toolOutputs = [];

      for (const toolCall of toolCalls) {
        const functionName = toolCall.function.name;

        console.log(`This question requires us to call a function: ${functionName}`);

        const args = JSON.parse(toolCall.function.arguments);
        //console.log(args);
        args.accessToken = accessToken;

        const argsArray = Object.keys(args).map((key) => args[key]);

        // Dynamically call the function with arguments
        const output5 = await window[functionName].apply(null, [args]);

        //console.log(`Output: ${output5}`);
        //console.log(output5);

        toolOutputs.push({
          tool_call_id: toolCall.id,
          output: `${output5}`,
        });
      }
      // Submit tool outputs
      await openai.beta.threads.runs.submitToolOutputs(thread_id, run.id, { tool_outputs: toolOutputs });
      continue; // Continue polling for the final response
    }

    // Check for failed, cancelled, or expired status
    if (["failed", "cancelled", "expired"].includes(runstatus.status)) {
      console.log(`Run status is '${runstatus.status}'. Unable to complete the request.`);
      break; // Exit the loop if the status indicates a failure or cancellation
    }
  }

  // Get the last assistant message from the messages array
  const messages = await openai.beta.threads.messages.list(thread_id);

  // Find the last message for the current run
  const lastMessageForRun = messages.data
    .filter((message) => message.run_id === run.id && message.role === "assistant")
    .pop();

  // If an assistant message is found, console.log() it
  if (lastMessageForRun) {
    console.log(`${lastMessageForRun.content[0].text.value} \n`);
  } else if (!["failed", "cancelled", "expired"].includes(runstatus.status)) {
    console.log("No response received from the assistant.");
  }

  output.textContent = lastMessageForRun.content[0].text.value;
  const pElement = document.createElement("p");
  pElement.textContent = inputElement.value;
  pElement.addEventListener("click", () => changeInput(pElement.textContent));
  history.append(pElement);
}

window.getAllMessages = getAllMessages;
