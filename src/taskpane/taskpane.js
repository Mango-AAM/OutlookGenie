/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//import { get } from "core-js/core/dict"
const axios = require('axios');

/* global document, Office */

const API_KEY = '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'
const submitButton = document.getElementById("submit")
const output = document.getElementById("output")
const inputElement = document.querySelector('input')
const history = document.querySelector('.history')
const buttonElement = document.querySelector('button')
const conversationHistory = [{role: "system", content: "You are a helpful assistant."}]

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
      if (result.status === "succeeded") {
        const accessToken = result.value;
        submitButton.addEventListener('click',getMessage);
        buttonElement.addEventListener('click', clearInput);
        console.log(accessToken);

        const restHost = Office.context.mailbox.restUrl;
        const searchKeyword = 'kill';
        const getMessageUrl = Office.context.mailbox.restUrl + `/v2.0/me/messages?$search=${encodeURIComponent(searchKeyword)}`;
        getAllMessages(accessToken, getMessageUrl, searchKeyword);
        console.log("Hahahaha");

      } else {
        console.error(result.error);
      }
    });
  }
});

function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  }
}

function changeInput(value){
  inputElement.value = value
}

function clearInput(){
  inputElement.value = ''
}


function getAllMessages(accessToken, getMessageUrl, searchKeyword) {
  console.log("Hahahaha");
  //const itemId = getItemRestId();
  //console.log(itemId);

  axios.get(getMessageUrl, {
    headers: {
        'Authorization': 'Bearer ' + accessToken
    }
  })
  .then(async function (response) {
    console.log(response.data);
    console.log(response.data.value.length);
    if (response.data.value && response.data.value.length > 0) {
      // Loop through each message in the result
      response.data.value.forEach(async function(message) {
        const subject = await message.Subject;
        console.log(subject);
        const body = await message.BodyPreview;
        console.log(`Body: ${body}`);
      });
    }
    if (response.data["@odata.nextLink"]) {
      await getAllMessages(accessToken, response.data["@odata.nextLink"]);
  }})
  .catch(function (error) {
      console.error('Error from Axios:', error);
  });
}


export async function getMessage(){
  console.log("Clicked");
  conversationHistory.push({role: "user", content: inputElement.value})
  const options = {
      method: 'POST',
      headers: {
          'Authorization': `Bearer ${API_KEY}`,
          'Content-Type' : 'application/json'
      },
      body: JSON.stringify({
          model: "gpt-3.5-turbo",
          messages: conversationHistory,
          temperature: 0.7,
          max_tokens: 100
        })
          
  }
  try{
      const response = await fetch('https://api.openai.com/v1/chat/completions',options)
      const data = await response.json()
      console.log(options)
      console.log("Hahahaha")
      console.log(data)
      output.textContent = data.choices[0].message.content
      conversationHistory.push(data.choices[0].message)
      const pElement = document.createElement('p')
      pElement.textContent = inputElement.value
      pElement.addEventListener('click',()=>changeInput(pElement.textContent))
      history.append(pElement)

  }catch(error){
      console.error(error);
  }
}

export async function run() {
  // Get a reference to the current message
  const item = Office.context.mailbox.item;

  // Write message property value to the task pane
  //document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  /**
   * Insert your Outlook code here
   */
}
