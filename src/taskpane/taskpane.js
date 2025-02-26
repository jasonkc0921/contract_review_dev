/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
(function() {
  'use strict';

  let reviewItems = [];
  let currentItemIndex = 0;
  let dialogPromise = null;
  let isReviewingSelection = false;
  let selectedRange = null;
  let selectionTrackingInterval = null;

  // Initialize the add-in
  Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
      // Initialize UI event handlers
      document.getElementById('start-review').onclick = startReview;
      document.getElementById('review-selection').onclick = reviewSelection;
      
      // Start monitoring for selections
      startSelectionTracking();
    }
  });
  
  // Track selections in the document to provide feedback to the user
  function startSelectionTracking() {
    // Clear any existing interval
    if (selectionTrackingInterval) {
      clearInterval(selectionTrackingInterval);
    }
    
    // Set up new interval to check for selection every second
    selectionTrackingInterval = setInterval(checkForSelection, 1000);
    
    // Initial check
    checkForSelection();
  }
  
  // Check if text is selected and update the status message
  async function checkForSelection() {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        
        const selectionStatus = document.getElementById('selection-status');
        
        if (!selection.text || selection.text.trim() === '') {
          selectionStatus.textContent = "No text selected. Select text to review a specific section.";
        } else {
          // Truncate if the selection is too long for display
          let displayText = selection.text;
          if (displayText.length > 50) {
            displayText = displayText.substring(0, 50) + "...";
          }
          
          selectionStatus.textContent = `Selection: "${displayText}"`;
        }
      });
    } catch (error) {
      console.error("Error checking selection:", error);
    }
  }

  // Check if text is selected and review it
  async function reviewSelection() {
    try {
      updateStatus("Checking for selected text...");
      
      await Word.run(async (context) => {
        // Get the current selection
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        
        if (!selection.text || selection.text.trim() === '') {
          updateStatus("No text is selected. Please select text to review.");
          return;
        }
        
        // Store the range for later use when applying changes
        selectedRange = selection;
        isReviewingSelection = true;
        
        updateStatus("Sending selected text to AI for review...");
        
        // Get the text from the selection
        const selectedText = selection.text;
        
        // Send to OpenAI API
        const recommendations = await getAIRecommendations(selectedText);
        
        if (recommendations && recommendations.length > 0) {
          reviewItems = recommendations;
          currentItemIndex = 0;
          
          // Show count in the task pane
          updateStatus(`Received ${recommendations.length} recommendations for selected text. Starting review...`);
          
          // Launch the dialog
          showReviewDialog();
        } else {
          updateStatus("No recommendations received from AI for the selected text.");
          isReviewingSelection = false;
          selectedRange = null;
        }
      });
    } catch (error) {
      updateStatus("Error: " + error.message);
      console.error(error);
      isReviewingSelection = false;
      selectedRange = null;
    }
  }

  // Fetch document text and send to OpenAI
  async function startReview() {
    updateStatus("Reading document...");
    isReviewingSelection = false;
    selectedRange = null;
    
    try {
      // Get the full text of the document
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        
        const documentText = body.text;
        
        updateStatus("Sending to AI for review...");
        
        // Send to OpenAI API
        const recommendations = await getAIRecommendations(documentText);
        
        if (recommendations && recommendations.length > 0) {
          reviewItems = recommendations;
          currentItemIndex = 0;
          
          // Show count in the task pane
          updateStatus(`Received ${recommendations.length} recommendations from AI. Starting review...`);
          
          // Launch the dialog
          showReviewDialog();
        } else {
          updateStatus("No recommendations received from AI.");
        }
      });
    } catch (error) {
      updateStatus("Error: " + error.message);
      console.error(error);
    }
  }

  // Function to call OpenAI API
  async function getAIRecommendations(text) {
    try {
      updateStatus("Requesting AI recommendations...");

      const openaiKey = window.ENV?.OPENAI_API_KEY;
    
      if (!openaiKey) {
        throw new Error("API key not found. Please check your environment configuration.");
      }
      
      const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${openaiKey}`
        },
        body: JSON.stringify({
          model: "gpt-4",
          messages: [
            { role: "system", content: "You are a labor law advisor who is specialized in labor law in Malaysia." },
            { role: "user", content: `please review the attached employment contract and suggest sections that need to be amended so that it conforms to Malaysia employment law in following text, output them in Json object that consists of original text and recommended text; the key for original text is "original_text"., and key for recommended text is "recommended_text"; you should avoid capturing original texts that are title of a section, typically short wordings, less than 10 words, that ended with 2 new lines "\n\n" or colon or dash- "${text}"` }
          ],
          max_tokens: 2500
        })
      });

      if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
      }

      const data = await response.json();
      console.log("Raw OpenAI response:", data);

      const content = data.choices[0].message.content;
      console.log("Content from OpenAI:", content);

      // For selected text, adjust the prompt based on context
      if (isReviewingSelection) {
        // Try to parse as JSON first
        try {
          const jsonResponse = JSON.parse(content);
          
          // If parsing succeeds, process normally
          if (jsonResponse && typeof jsonResponse === 'object') {
            // Process as normal JSON
            const recommendations = [];
            
            // Check if it's an array of recommendations
            if (Array.isArray(jsonResponse)) {
              for (const item of jsonResponse) {
                if (item.original_text && item.recommended_text) {
                  recommendations.push({
                    original: item.original_text,
                    recommended: item.recommended_text
                  });
                }
              }
            } 
            // Check if it's a single recommendation
            else if (jsonResponse.original_text && jsonResponse.recommended_text) {
              recommendations.push({
                original: jsonResponse.original_text,
                recommended: jsonResponse.recommended_text
              });
            }
            
            // If we got recommendations, return them
            if (recommendations.length > 0) {
              return recommendations;
            }
          }
        } catch (jsonError) {
          console.log("Response is not valid JSON, continuing with text parsing");
        }
      }

      // Direct extraction approach - don't try to parse complete JSON
      const recommendations = [];
      
      // Extract original and recommended text content using regex
      // This regex looks for "original_text": and "recommended_text": patterns
      const originalRegex = /"original_text"\s*:\s*"((\\"|[^"])*?)"/g;
      const recommendedRegex = /"recommended_text"\s*:\s*"((\\"|[^"])*?)"/g;
      
      const originals = [];
      const recommendeds = [];
      
      // Extract all original_text values
      let match;
      while ((match = originalRegex.exec(content)) !== null) {
        // Replace escaped quotes and newlines
        const cleanedText = match[1]
          .replace(/\\"/g, '"')
          .replace(/\\n/g, '\n')
          .replace(/\\\\/g, '\\');
        originals.push(cleanedText);
      }
      
      // Extract all recommended_text values
      while ((match = recommendedRegex.exec(content)) !== null) {
        // Replace escaped quotes and newlines
        const cleanedText = match[1]
          .replace(/\\"/g, '"')
          .replace(/\\n/g, '\n')
          .replace(/\\\\/g, '\\');
        recommendeds.push(cleanedText);
      }
      
      console.log("Extracted originals:", originals);
      console.log("Extracted recommendeds:", recommendeds);
      
      // Match originals with recommendeds
      // If counts match, pair them in order
      if (originals.length === recommendeds.length) {
        for (let i = 0; i < originals.length; i++) {
          recommendations.push({
            original: originals[i],
            recommended: recommendeds[i]
          });
        }
      } else {
        // If counts don't match, try alternative extraction: look for blocks
        console.log("Counts don't match, trying alternative extraction");
        
        // Try to identify blocks of text that contain both original and recommended
        const blocks = content.split(/(?=\{\s*"original_text")/);
        
        for (const block of blocks) {
          if (!block.includes("original_text") || !block.includes("recommended_text")) {
            continue;
          }
          
          const originalMatch = block.match(/"original_text"\s*:\s*"((\\"|[^"])*?)"/);
          const recommendedMatch = block.match(/"recommended_text"\s*:\s*"((\\"|[^"])*?)"/);
          
          if (originalMatch && recommendedMatch) {
            const originalText = originalMatch[1]
              .replace(/\\"/g, '"')
              .replace(/\\n/g, '\n')
              .replace(/\\\\/g, '\\');
              
            const recommendedText = recommendedMatch[1]
              .replace(/\\"/g, '"')
              .replace(/\\n/g, '\n')
              .replace(/\\\\/g, '\\');
              
            recommendations.push({
              original: originalText,
              recommended: recommendedText
            });
          }
        }
      }
      
      // Special case: If we're reviewing a selection and no recommendations found
      // but we have a direct response, treat the entire response as a single recommendation
      if (recommendations.length === 0 && isReviewingSelection) {
        console.log("Using direct response as recommendation for selected text");
        
        // For selected text, sometimes GPT just gives a direct updated version
        // In this case, use the selected text as original and GPT response as recommendation
        if (text && content) {
          recommendations.push({
            original: text,
            recommended: content
          });
        }
      }
      
      // If all attempts fail, try a simpler approach with line-by-line parsing
      if (recommendations.length === 0) {
        console.log("Block approach failed, trying line-by-line");
        
        const lines = content.split('\n');
        let currentOriginal = null;
        
        for (let i = 0; i < lines.length; i++) {
          const line = lines[i];
          
          if (line.includes('"original_text"')) {
            // Extract original text
            const textMatch = line.match(/"original_text"\s*:\s*"((\\"|[^"])*?)"/);
            if (textMatch) {
              currentOriginal = textMatch[1]
                .replace(/\\"/g, '"')
                .replace(/\\n/g, '\n')
                .replace(/\\\\/g, '\\');
            }
          } else if (line.includes('"recommended_text"') && currentOriginal !== null) {
            // Extract recommended text and create a pair
            const textMatch = line.match(/"recommended_text"\s*:\s*"((\\"|[^"])*?)"/);
            if (textMatch) {
              const recommendedText = textMatch[1]
                .replace(/\\"/g, '"')
                .replace(/\\n/g, '\n')
                .replace(/\\\\/g, '\\');
                
              recommendations.push({
                original: currentOriginal,
                recommended: recommendedText
              });
              
              currentOriginal = null;
            }
          }
        }
      }
      
      console.log("Final processed recommendations:", recommendations);
      return recommendations;
    } catch (error) {
      console.error('Error calling OpenAI:', error);
      updateStatus('Failed to get AI recommendations: ' + error.message);
      return [];
    }
  }
  
  // Show the review dialog using Office Dialog API
  function showReviewDialog() {
    // Prepare the data to pass to the dialog
    const currentItem = reviewItems[currentItemIndex];
    
    // Encode data as URL parameters
    const dialogUrl = new URL('./dialog.html', window.location.href);
    dialogUrl.searchParams.set('original', encodeURIComponent(currentItem.original));
    dialogUrl.searchParams.set('recommended', encodeURIComponent(currentItem.recommended));
    dialogUrl.searchParams.set('itemIndex', currentItemIndex);
    dialogUrl.searchParams.set('totalItems', reviewItems.length);
    // Add a flag for selection mode
    dialogUrl.searchParams.set('isSelection', isReviewingSelection);
    
    console.log("Opening dialog with URL:", dialogUrl.href);
    
    // Open the dialog with the data already in the URL
    Office.context.ui.displayDialogAsync(
        dialogUrl.href,
        { width: 80, height: 60, displayInIframe: true },
        function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                dialogPromise = result.value;
                
                // Add message handler for button actions only
                dialogPromise.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
                    try {
                        const message = JSON.parse(arg.message);
                        processDialogMessage(arg);
                    } catch (error) {
                        console.error("Error processing dialog message:", error);
                    }
                });
                
                dialogPromise.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
            } else {
                console.error("Error opening dialog:", result.error);
            }
        }
    );
  }

  function sendDataToDialog() {
    if (dialogPromise && reviewItems.length > 0) {
        // Create a simple, flat object structure
        const dataToSend = {
            messageType: 'reviewData',
            itemIndex: currentItemIndex,
            totalItems: reviewItems.length,
            isSelection: isReviewingSelection,
            currentItem: {
                original: String(reviewItems[currentItemIndex].original || ''),
                recommended: String(reviewItems[currentItemIndex].recommended || '')
            }
        };
        
        // Log the exact string we're sending
        const messageString = JSON.stringify(dataToSend);
        console.log("Sending exact string:", messageString);
        
        try {
            dialogPromise.messageChild(messageString);
            console.log("Message sent successfully");
        } catch (error) {
            console.error("Error sending message:", error);
        }
    }
  }

  // Process messages received from the dialog
  function processDialogMessage(arg) {
    try {
      const message = JSON.parse(arg.message);
      console.log("Message received from dialog:", message);
      
      switch(message.action) {
        case 'approve':
          // Use the edited text if it exists
          if (message.editedText) {
            // Create a copy of the current item with the edited text
            const editedItem = {
              original: reviewItems[currentItemIndex].original,
              recommended: message.editedText
            };
            // Replace the current item with the edited version
            reviewItems[currentItemIndex] = editedItem;
          }
          handleApprove();
          break;
        case 'next':
          handleNext();
          break;
        case 'cancel':
          handleCancel();
          break;
        case 'addComment':
          handleAddComment(message.commentText);
          break;
        case 'ready':
          // Dialog is ready to receive data
          sendDataToDialog();
          break;
        case 'requestNewSuggestion':
          // Request a new suggestion from OpenAI
          handleNewSuggestionRequest(message.originalText);
          break;
      }
    } catch (error) {
      console.error("Error processing dialog message:", error);
    }
  }
  
  // Process dialog events (e.g., dialog closed)
  function processDialogEvent(arg) {
    if (arg.error === 12006) {
      // Dialog was closed
      dialogPromise = null;
      updateStatus("Review window closed.");
    }
  }
  
  // Handle request for a new suggestion
  async function handleNewSuggestionRequest(originalText) {
    try {
      console.log("Requesting new suggestion for text:", originalText);
      updateStatus("Requesting new suggestion from AI...");
      
      // Make a direct call to OpenAI for a single suggestion
      const newSuggestion = await getNewSuggestionFromAI(originalText);
      
      if (newSuggestion && dialogPromise) {
        console.log("Received new suggestion from AI:", newSuggestion);
        
        // Send the new suggestion back to the dialog
        try {
          const messageToSend = JSON.stringify({
            messageType: 'newSuggestion',
            newText: newSuggestion
          });
          
          console.log("Sending new suggestion to dialog:", messageToSend);
          dialogPromise.messageChild(messageToSend);
          console.log("New suggestion sent to dialog");
          
          updateStatus("New suggestion received from AI.");
        } catch (error) {
          console.error("Error sending new suggestion to dialog:", error);
          updateStatus("Error sending new suggestion to dialog: " + error.message);
        }
      } else {
        console.error("Failed to get new suggestion or dialog is closed");
        updateStatus("Failed to get new suggestion.");
        
        if (dialogPromise) {
          try {
            dialogPromise.messageChild(JSON.stringify({
              messageType: 'newSuggestion',
              error: "Failed to generate a new suggestion."
            }));
          } catch (error) {
            console.error("Error sending failure message to dialog:", error);
          }
        }
      }
    } catch (error) {
      console.error("Error getting new suggestion:", error);
      updateStatus("Error: " + error.message);
      
      if (dialogPromise) {
        try {
          dialogPromise.messageChild(JSON.stringify({
            messageType: 'newSuggestion',
            error: error.message
          }));
        } catch (commError) {
          console.error("Error sending error message to dialog:", commError);
        }
      }
    }
  }
  
  // Function to get a new suggestion from OpenAI
  async function getNewSuggestionFromAI(originalText) {
    try {
      console.log("Calling OpenAI API for new suggestion");

      const openaiKey = window.ENV?.OPENAI_API_KEY;
    
      if (!openaiKey) {
        throw new Error("API key not found. Please check your environment configuration.");
      }
      
      const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${openaiKey}`
        },
        body: JSON.stringify({
          model: "gpt-4",
          messages: [
            { 
              role: "system", 
              content: "You are a labor law advisor specialized in Malaysian labor law. You're being asked to revise a specific clause from an employment contract to ensure it conforms to Malaysian employment law. Provide ONLY the revised text without any explanations or preamble."
            },
            { 
              role: "user", 
              content: `Please review and revise the following employment contract clause to fully comply with Malaysian labor law. Return ONLY the revised text: "${originalText}"`
            }
          ],
          max_tokens: 1000
        })
      });

      if (!response.ok) {
        throw new Error(`API request failed with status ${response.status}`);
      }

      const data = await response.json();
      console.log("Raw OpenAI response for new suggestion:", data);

      // Extract the suggested text
      const suggestedText = data.choices[0].message.content.trim();
      console.log("Extracted suggestion:", suggestedText);
      return suggestedText;
    } catch (error) {
      console.error('Error calling OpenAI for new suggestion:', error);
      throw error;
    }
  }
  
  // Handle approve action from dialog
  async function handleApprove() {
    try {
        console.log(`Attempting to approve item at index: ${currentItemIndex}`);
        console.log(`Total review items: ${reviewItems.length}`);
        
        // Get the current item to approve
        let currentItem = reviewItems[currentItemIndex];
        let originalText = currentItem.original;
        let recommendedText = currentItem.recommended;
        
        console.log(`Current Original Text (first 100 chars): ${originalText.substring(0, 100)}...`);
        console.log(`Current Recommended Text (first 100 chars): ${recommendedText.substring(0, 100)}...`);
        
        await Word.run(async (context) => {
            // If we're reviewing a selection, replace the selection directly
            if (isReviewingSelection && selectedRange) {
                console.log("Updating selected text");
                
                // Get a fresh reference to the selected range
                const selection = context.document.getSelection();
                
                // Replace the selection with the recommended text
                recommendedText = recommendedText.replace(/\\n/g, '\n');
                selection.insertText(recommendedText, Word.InsertLocation.replace);
                await context.sync();
                
                // Show comment dialog
                if (dialogPromise) {
                    dialogPromise.messageChild(JSON.stringify({
                        messageType: 'promptComment'
                    }));
                }
                
                return;
            }
            
            // For full document review, search for the text to replace
            // Improved text normalization for search words
            const getFirst15Words = (text) => {
                const cleanText = text
                    .replace(/\n/g, ' ')
                    .replace(/\\+/g, '')
                    .replace(/\s+/g, ' ')
                    .trim();
                
                const words = cleanText.split(' ')
                    .filter(word => word.length > 0)
                    .slice(0, 15);
                
                console.log("First 15 individual words:", words);
                return words;
            };

            const searchWords = getFirst15Words(originalText);
            console.log("Search words:", searchWords);
            
            // Load paragraphs with limited formatting properties that are well-supported
            const paragraphs = context.document.body.paragraphs;
            context.load(paragraphs, ['text', 'font', 'style']);
            await context.sync();

            let found = false;
            for (let i = 0; i < paragraphs.items.length; i++) {
                // Normalize paragraph text
                let combinedText = paragraphs.items[i].text;
                const cleanParaText = combinedText
                    .replace(/\n/g, ' ')
                    .replace(/\s+/g, ' ')
                    .trim();
                
                const paraWords = cleanParaText.split(' ').filter(word => word.length > 0);
                
                let allWordsFound = true;
                let lastIndex = -1;
                
                // Check if all search words are found in order
                for (const word of searchWords) {
                    const index = paraWords.findIndex((paraWord, idx) => 
                        idx > lastIndex && 
                        paraWord.toLowerCase().trim() === word.toLowerCase().trim()
                    );
                    
                    if (index === -1) {
                        allWordsFound = false;
                        break;
                    }
                    lastIndex = index;
                }

                if (allWordsFound) {
                    console.log(`Found matching text starting in paragraph ${i}`);
                    
                    // Normalize original text for comparison
                    const normalizedOriginal = originalText
                        .replace(/\n/g, ' ')
                        .replace(/\s+/g, ' ')
                        .trim();

                    // Find all paragraphs that contain parts of the original text
                    let paragraphsToRemove = [];
                    let remainingText = normalizedOriginal;
                    let j = i;
                    
                    while (remainingText.length > 0 && j < paragraphs.items.length) {
                        const normalizedParaText = paragraphs.items[j].text
                            .replace(/\n/g, ' ')
                            .replace(/\s+/g, ' ')
                            .trim();
                        
                        if (normalizedOriginal.includes(normalizedParaText) ||
                            normalizedParaText.includes(remainingText.slice(0, 50))) {
                            paragraphsToRemove.push(j);
                            remainingText = remainingText.replace(normalizedParaText, '').trim();
                        } else {
                            break;
                        }
                        j++;
                    }

                    // An alternative approach: instead of modifying the paragraph,
                    // preserve the exact formatting by getting its HTML
                    const originalParagraph = paragraphs.items[i];
                    
                    // Instead of trying to manipulate specific formatting properties directly,
                    // we'll use a range-based approach with insertHtml which better preserves formatting
                    
                    // First, create a range from the found paragraph
                    const range = originalParagraph.getRange();
                    
                    // Replace text with new content
                    recommendedText = recommendedText.replace(/\\n/g, '\n');
                    range.insertText(recommendedText, Word.InsertLocation.replace);
                    await context.sync();
                    
                    // Delete additional paragraphs that contained parts of original text
                    for (let idx = 1; idx < paragraphsToRemove.length; idx++) {
                        paragraphs.items[paragraphsToRemove[idx]].delete();
                    }
                    await context.sync();

                    found = true;

                    // Show comment dialog
                    if (dialogPromise) {
                        dialogPromise.messageChild(JSON.stringify({
                            messageType: 'promptComment'
                        }));
                    }
                    break;
                }
            }

            if (!found && !isReviewingSelection) {
                console.error(`Could not find matching text for item ${currentItemIndex}`);
                updateStatus(`Error: Could not find the text to replace for item ${currentItemIndex}`);
                
                // Even if text wasn't found, still show comment dialog or move to next
                if (dialogPromise) {
                    dialogPromise.messageChild(JSON.stringify({
                        messageType: 'promptComment'
                    }));
                }
            }
        });
    } catch (error) {
        console.error("Error in handleApprove:", error);
        updateStatus(`Error during approval: ${error.message}`);
        
        // If error occurs, still allow adding comment or moving to next
        if (dialogPromise) {
            dialogPromise.messageChild(JSON.stringify({
                messageType: 'promptComment'
            }));
        }
    }
  }

  async function handleAddComment(commentText) {
    if (!commentText) {
        handleNext();
        return;
    }
    
    try {
        await Word.run(async (context) => {
            let target;
            
            // Different approach for selected text vs full document review
            if (isReviewingSelection) {
                // For selection, add comment to the current selection
                target = context.document.getSelection();
            } else {
                // For full document, look for the recently inserted text
                const searchPhrase = reviewItems[currentItemIndex].recommended.substring(0, 20);
                const searchResults = context.document.body.search(searchPhrase, {matchCase: false});
                context.load(searchResults);
                await context.sync();
                
                if (searchResults.items.length > 0) {
                    target = searchResults.items[0];
                } else {
                    console.error("Could not find replaced text to add comment");
                    handleNext();
                    return;
                }
            }
            
            // Add the comment
            const comment = target.insertComment(commentText);
            context.load(comment);
            await context.sync();
            console.log("Comment added successfully");
            
            // Move to next item
            handleNext();
        });
    } catch (error) {
        console.error("Error adding comment:", error);
        handleNext();
    }
  }
  
  // Move to next item
  function handleNext() {
    if (currentItemIndex < reviewItems.length - 1) {
        currentItemIndex++;
        console.log(`Moving to next item ${currentItemIndex + 1} of ${reviewItems.length}`);
        
        // Make sure to send updated data to dialog
        setTimeout(() => {
            sendDataToDialog();
        }, 200); // Small delay to ensure processing completes
    } else {
        console.log("Review completed");
        completeReview();
    }
  }

  function sendDataToDialog() {
    if (dialogPromise && reviewItems.length > 0) {
        console.log(`Preparing dialog data for item ${currentItemIndex + 1}`);
        
        // For URL parameter approach, create new dialog with updated params
        const dialogUrl = new URL('./dialog.html', window.location.href);
        
        // Always get the current item from the current index
        const currentItem = reviewItems[currentItemIndex];
        dialogUrl.searchParams.set('original', encodeURIComponent(currentItem.original));
        dialogUrl.searchParams.set('recommended', encodeURIComponent(currentItem.recommended));
        dialogUrl.searchParams.set('itemIndex', currentItemIndex);
        dialogUrl.searchParams.set('totalItems', reviewItems.length);
        
        console.log("Opening dialog with updated URL:", dialogUrl.href);
        
        // Close existing dialog and reopen with new URL
        if (dialogPromise) {
            dialogPromise.close();
            dialogPromise = null;
            
            // Open new dialog with updated data
            setTimeout(() => {
                Office.context.ui.displayDialogAsync(
                    dialogUrl.href,
                    { width: 80, height: 60, displayInIframe: true },
                    function(result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            dialogPromise = result.value;
                            // Add handlers
                            dialogPromise.addEventHandler(Office.EventType.DialogMessageReceived, processDialogMessage);
                            dialogPromise.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
                        } else {
                            console.error("Error opening updated dialog:", result.error);
                            updateStatus("Error opening review dialog. Please try again.");
                        }
                    }
                );
            }, 300);
        }
    }
  }
  
  // Cancel the review process
  function handleCancel() {
    if (dialogPromise) {
      dialogPromise.close();
      dialogPromise = null;
    }
    updateStatus("Review canceled.");
    // Reset selection flags if reviewing a selection
    if (isReviewingSelection) {
      isReviewingSelection = false;
      selectedRange = null;
    }
  }

  // Complete the review process
  function completeReview() {
    if (dialogPromise) {
      dialogPromise.close();
      dialogPromise = null;
    }
    
    // Reset selection flags if reviewing a selection
    if (isReviewingSelection) {
      isReviewingSelection = false;
      selectedRange = null;
    }
    
    updateStatus("Review completed! All recommendations have been processed.");
  }

  // Update status message
  function updateStatus(message) {
    const statusElement = document.getElementById('status');
    if (statusElement) {
        statusElement.textContent = message;
    } else {
        console.error('Status element not found:', message);
    }
  }
})();