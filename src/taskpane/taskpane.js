/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

const API_KEY = ''
const submitButton = document.getElementById("submit")
const output = document.getElementById("output")
const inputElement = document.querySelector('input')
const history = document.querySelector('.history')
const buttonElement = document.querySelector('button')

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    submitButton.addEventListener('click',getMessage);
    buttonElement.addEventListener('click', clearInput);

  }
});

function changeInput(value){
  inputElement.value = value
}

function clearInput(){
  inputElement.value = ''
}


export async function getMessage(){
  console.log("Clicked");
  const options = {
      method: 'POST',
      headers: {
          'Authorization': `Bearer ${API_KEY}`,
          'Content-Type' : 'application/json'
      },
      body: JSON.stringify({
          model: "gpt-3.5-turbo",
          messages: [{role: "user", content: inputElement.value}],
          temperature: 0.7,
          max_tokens: 100
        })
          
  }
  try{
      const response = await fetch('https://api.openai.com/v1/chat/completions',options)
      const data = await response.json()
      console.log(data)
      output.textContent = data.choices[0].message.content
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
