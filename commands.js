/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

/**
 * Handles the button click event in compose mode to rewrite the email professionally.
 * @param event {Office.AddinCommands.Event}
 */
async function onComposeClick(event) {
    try {
        // Get the current message body
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const emailBody = result.value;
                // Call ChatGPT to rewrite the email professionally
                const rewrittenText = await callChatGPTApi(emailBody);
                // Replace the email body with the rewritten text
                Office.context.mailbox.item.body.setAsync(rewrittenText, { coercionType: Office.CoercionType.Text });
            } else {
                console.error("Failed to get email body:", result.error);
            }
        });
    } catch (error) {
        console.error("Error in rewriting email:", error);
    }
    event.completed(); // Signal that the function is complete
}

// Function to call the OpenAI API and rewrite the email text
async function callChatGPTApi(text) {
    const apiKey = 'sk-nqOIhwl1gY9Uf0CUTaU2aAeVc90GY40BLQVQLtY-AdT3BlbkFJQHQD95oKXg0OCpYZawjHMVaigACSUW0wf9uIreeVwA'; // Replace with your OpenAI API key
    const apiUrl = 'https://api.openai.com/v1/chat/completions';

    const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
            model: "gpt-3.5-turbo",
            messages: [
                { role: "system", content: "Make this email sound professional:" },
                { role: "user", content: text }
            ]
        })
    });

    if (!response.ok) {
        throw new Error(`Error calling OpenAI API: ${response.statusText}`);
    }

    const data = await response.json();
    return data.choices[0].message.content;
}

// Register the compose function with Office.
Office.actions.associate("onComposeClick", onComposeClick);
