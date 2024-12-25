/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// eslint-disable-next-line no-undef
const fetch = require("node-fetch");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// Function to call the OpenAI API
async function fetchOpenAIResponse() {
  const apiKey = ""; // Replace this with your actual OpenAI API key
  const url = "https://api.openai.com/v1/completions";

  const prompt = "Write something that will be inserted into Excel."; // You can customize this prompt as needed

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "text-davinci-003", // You can use different models if needed
      prompt: prompt,
      max_tokens: 50,
    }),
  });

  const data = await response.json();
  return data.choices[0].text.trim();
}

// Function to write the OpenAI response into Excel
async function writeToExcel(responseText) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1:A10"); // Adjust the range as needed

    const responseArray = responseText.split("\n"); // In case the response has multiple lines

    for (let i = 0; i < responseArray.length && i < 10; i++) {
      range.getCell(i, 0).values = [[responseArray[i]]]; // Write each line to consecutive rows
    }

    await context.sync();
  });
}

// Event listener for the button click
document.getElementById("fetchOpenAIResponse").addEventListener("click", async () => {
  try {
    const openAIResponse = await fetchOpenAIResponse();
    await writeToExcel(openAIResponse);
  } catch (error) {
    console.error("Error:", error);
  }
});
