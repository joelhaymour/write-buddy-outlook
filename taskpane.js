/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.addEventListener("DOMContentLoaded", initializeAddIn);
  }
});

let currentSelectedText = '';
let currentRange = null;
let isRewriting = false;

// Initialize the add-in
function initializeAddIn() {
  // Load settings
  loadSettings();
  
  // Get selected text button
  document.getElementById('get-selected-text').addEventListener('click', getSelectedText);
  
  // Option buttons
  document.getElementById('btn-professional').addEventListener('click', () => handleRewrite('professional'));
  document.getElementById('btn-factcheck').addEventListener('click', () => handleRewrite('factcheck'));
  document.getElementById('btn-persuasive').addEventListener('click', () => handleRewrite('persuasive'));
  document.getElementById('btn-translate').addEventListener('click', handleTranslate);
  
  // Insert and Clear buttons
  document.getElementById('btn-insert').addEventListener('click', insertText);
  document.getElementById('btn-cancel').addEventListener('click', clearText);
  
  // Settings button
  document.getElementById('btn-settings').addEventListener('click', openSettings);
  
  // Feature toggles
  document.getElementById('toggle-professional').addEventListener('change', saveFeatureToggles);
  document.getElementById('toggle-factcheck').addEventListener('change', saveFeatureToggles);
  document.getElementById('toggle-persuasive').addEventListener('change', saveFeatureToggles);
  document.getElementById('toggle-translate').addEventListener('change', saveFeatureToggles);
  
  // Update button visibility based on toggles
  updateButtonVisibility();
}

// Get selected text from Outlook
async function getSelectedText() {
  try {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      { asyncContext: "getSelectedText" },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Try to get selection using Office.js
          Office.context.mailbox.item.body.getSelectedDataAsync(
            Office.CoercionType.Text,
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const selectedText = asyncResult.value.data;
                if (selectedText && selectedText.trim()) {
                  currentSelectedText = selectedText.trim();
                  document.getElementById('original-text').value = formatTextForDisplay(currentSelectedText);
                  showStatus('Text selected successfully!', 'success');
                } else {
                  showStatus('Please select some text in the email body first.', 'warning');
                }
              } else {
                // Fallback: prompt user to select text
                showStatus('Please select text in the email body, then click this button again.', 'info');
              }
            }
          );
        } else {
          showStatus('Error accessing email body.', 'error');
        }
      }
    );
  } catch (error) {
    console.error('Error getting selected text:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

// Format text for display
function formatTextForDisplay(text) {
  let formatted = text.replace(/\n{3,}/g, '\n\n');
  formatted = formatted.split('\n').map(line => line.trim()).join('\n');
  return formatted.trim();
}

// Handle rewrite
async function handleRewrite(style) {
  const originalText = document.getElementById('original-text').value.trim();
  
  if (!originalText) {
    showStatus('Please get selected text first.', 'warning');
    return;
  }
  
  isRewriting = true;
  const rewrittenTextarea = document.getElementById('rewritten-text');
  const insertBtn = document.getElementById('btn-insert');
  
  // Disable buttons
  setButtonsEnabled(false);
  rewrittenTextarea.value = 'Rewriting...';
  
  try {
    // Get API key
    const apiKey = await getApiKey();
    if (!apiKey) {
      throw new Error('API key not set. Please configure it in settings.');
    }
    
    // Call rewrite API
    const response = await callRewriteAPI(originalText, style, apiKey);
    
    if (response && response.success && response.rewrittenText) {
      const formattedText = formatTextForDisplay(response.rewrittenText);
      rewrittenTextarea.value = formattedText;
      insertBtn.disabled = false;
      showStatus('Text rewritten successfully!', 'success');
    } else {
      throw new Error(response?.error || 'Rewrite failed');
    }
  } catch (error) {
    console.error('Rewrite error:', error);
    rewrittenTextarea.value = `Error: ${error.message}`;
    showStatus('Error: ' + error.message, 'error');
  } finally {
    setButtonsEnabled(true);
    isRewriting = false;
  }
}

// Handle translate
async function handleTranslate() {
  const originalText = document.getElementById('original-text').value.trim();
  
  if (!originalText) {
    showStatus('Please get selected text first.', 'warning');
    return;
  }
  
  // Show language selector
  showLanguageSelector();
}

// Show language selector
function showLanguageSelector() {
  const languages = [
    { code: 'es', name: 'Spanish' },
    { code: 'fr', name: 'French' },
    { code: 'de', name: 'German' },
    { code: 'it', name: 'Italian' },
    { code: 'pt', name: 'Portuguese' },
    { code: 'ru', name: 'Russian' },
    { code: 'ja', name: 'Japanese' },
    { code: 'ko', name: 'Korean' },
    { code: 'zh', name: 'Chinese (Simplified)' },
    { code: 'ar', name: 'Arabic' },
    { code: 'hi', name: 'Hindi' },
    { code: 'nl', name: 'Dutch' },
    { code: 'pl', name: 'Polish' },
    { code: 'tr', name: 'Turkish' },
    { code: 'sv', name: 'Swedish' },
    { code: 'da', name: 'Danish' },
    { code: 'no', name: 'Norwegian' },
    { code: 'fi', name: 'Finnish' },
    { code: 'cs', name: 'Czech' },
    { code: 'el', name: 'Greek' },
    { code: 'he', name: 'Hebrew' },
    { code: 'th', name: 'Thai' },
    { code: 'vi', name: 'Vietnamese' },
    { code: 'id', name: 'Indonesian' }
  ];
  
  // Create language selector modal
  const modal = document.createElement('div');
  modal.id = 'lang-selector-modal';
  modal.className = 'lang-selector-modal';
  modal.innerHTML = `
    <div class="lang-selector-content">
      <h3>Select Language</h3>
      <input type="text" id="lang-search" placeholder="Search language..." autocomplete="off">
      <div id="lang-list" class="lang-list"></div>
      <button class="lang-close-btn">Cancel</button>
    </div>
  `;
  
  const langList = modal.querySelector('#lang-list');
  const searchInput = modal.querySelector('#lang-search');
  
  function renderLanguages(filter = '') {
    const filtered = languages.filter(lang => 
      lang.name.toLowerCase().includes(filter.toLowerCase())
    );
    
    langList.innerHTML = filtered.map(lang => 
      `<div class="lang-item" data-code="${lang.code}">${lang.name}</div>`
    ).join('');
    
    langList.querySelectorAll('.lang-item').forEach(item => {
      item.addEventListener('click', async () => {
        const langCode = item.getAttribute('data-code');
        const langName = item.textContent;
        document.body.removeChild(modal);
        await translateText(langCode, langName);
      });
    });
  }
  
  searchInput.addEventListener('input', (e) => {
    renderLanguages(e.target.value);
  });
  
  modal.querySelector('.lang-close-btn').addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  renderLanguages();
  document.body.appendChild(modal);
  searchInput.focus();
}

// Translate text
async function translateText(langCode, langName) {
  const originalText = document.getElementById('original-text').value.trim();
  
  if (!originalText) {
    showStatus('Please get selected text first.', 'warning');
    return;
  }
  
  isRewriting = true;
  const rewrittenTextarea = document.getElementById('rewritten-text');
  const insertBtn = document.getElementById('btn-insert');
  
  setButtonsEnabled(false);
  rewrittenTextarea.value = `Translating to ${langName}...`;
  
  try {
    const apiKey = await getApiKey();
    if (!apiKey) {
      throw new Error('API key not set. Please configure it in settings.');
    }
    
    const response = await callTranslateAPI(originalText, langCode, apiKey);
    
    if (response && response.success && response.translatedText) {
      const formattedText = formatTextForDisplay(response.translatedText);
      rewrittenTextarea.value = formattedText;
      insertBtn.disabled = false;
      showStatus('Text translated successfully!', 'success');
    } else {
      throw new Error(response?.error || 'Translation failed');
    }
  } catch (error) {
    console.error('Translation error:', error);
    rewrittenTextarea.value = `Error: ${error.message}`;
    showStatus('Error: ' + error.message, 'error');
  } finally {
    setButtonsEnabled(true);
    isRewriting = false;
  }
}

// Insert text into email
async function insertText() {
  const rewrittenText = document.getElementById('rewritten-text').value.trim();
  
  if (!rewrittenText) {
    showStatus('No text to insert.', 'warning');
    return;
  }
  
  try {
    // Format the text for insertion
    const formattedText = formatTextForDisplay(rewrittenText);
    
    // Try to replace selected text first
    Office.context.mailbox.item.body.getSelectedDataAsync(
      Office.CoercionType.Text,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value.data) {
          // We have a selection - replace it
          Office.context.mailbox.item.body.setSelectedDataAsync(
            formattedText,
            { coercionType: Office.CoercionType.Text },
            (setResult) => {
              if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                showStatus('Text inserted successfully!', 'success');
                document.getElementById('rewritten-text').value = '';
                document.getElementById('btn-insert').disabled = true;
              } else {
                // Fallback: try HTML replacement
                replaceTextInBody(formattedText);
              }
            }
          );
        } else {
          // No selection - append to end or insert at cursor
          insertTextAtCursor(formattedText);
        }
      }
    );
  } catch (error) {
    console.error('Error inserting text:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

// Replace text in body HTML (fallback method)
function replaceTextInBody(newText) {
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let bodyHtml = result.value;
        
        // Get selected text to find and replace
        Office.context.mailbox.item.body.getSelectedDataAsync(
          Office.CoercionType.Text,
          (selResult) => {
            if (selResult.status === Office.AsyncResultStatus.Succeeded && selResult.value.data) {
              const selectedText = selResult.value.data;
              // Escape HTML special characters in selected text for replacement
              const escapedSelected = selectedText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              const htmlNewText = newText.replace(/\n/g, '<br>');
              bodyHtml = bodyHtml.replace(new RegExp(escapedSelected, 'g'), htmlNewText);
              
              Office.context.mailbox.item.body.setAsync(
                bodyHtml,
                { coercionType: Office.CoercionType.Html },
                (setResult) => {
                  if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                    showStatus('Text inserted successfully!', 'success');
                    document.getElementById('rewritten-text').value = '';
                    document.getElementById('btn-insert').disabled = true;
                  } else {
                    showStatus('Error inserting text: ' + setResult.error.message, 'error');
                  }
                }
              );
            } else {
              insertTextAtCursor(newText);
            }
          }
        );
      } else {
        showStatus('Error accessing email body.', 'error');
      }
    }
  );
}

// Insert text at cursor position
function insertTextAtCursor(text) {
  // Get current cursor position and insert text
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let bodyHtml = result.value;
        const htmlText = text.replace(/\n/g, '<br>');
        
        // Try to get selection range
        Office.context.mailbox.item.body.getSelectedDataAsync(
          Office.CoercionType.Html,
          (selResult) => {
            if (selResult.status === Office.AsyncResultStatus.Succeeded && selResult.value.data) {
              // Replace selection
              const selectedHtml = selResult.value.data;
              bodyHtml = bodyHtml.replace(selectedHtml, htmlText);
            } else {
              // Append to end
              bodyHtml += '<br><br>' + htmlText;
            }
            
            Office.context.mailbox.item.body.setAsync(
              bodyHtml,
              { coercionType: Office.CoercionType.Html },
              (setResult) => {
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                  showStatus('Text inserted successfully!', 'success');
                  document.getElementById('rewritten-text').value = '';
                  document.getElementById('btn-insert').disabled = true;
                } else {
                  showStatus('Error inserting text: ' + setResult.error.message, 'error');
                }
              }
            );
          }
        );
      } else {
        showStatus('Error accessing email body.', 'error');
      }
    }
  );
}

// Clear text
function clearText() {
  document.getElementById('original-text').value = '';
  document.getElementById('rewritten-text').value = '';
  document.getElementById('btn-insert').disabled = true;
  currentSelectedText = '';
  showStatus('Cleared.', 'info');
}

// Get API key from storage
async function getApiKey() {
  return new Promise((resolve) => {
    try {
      const apiKey = Office.context.roamingSettings.get('openaiApiKey');
      resolve(apiKey || null);
    } catch (error) {
      console.error('Error getting API key:', error);
      resolve(null);
    }
  });
}

// Save API key
async function saveApiKey(apiKey) {
  return new Promise((resolve) => {
    try {
      Office.context.roamingSettings.set('openaiApiKey', apiKey);
      Office.context.roamingSettings.saveAsync((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded);
      });
    } catch (error) {
      console.error('Error saving API key:', error);
      resolve(false);
    }
  });
}

// Call rewrite API
async function callRewriteAPI(text, style, apiKey) {
  const prompt = getPromptForStyle(text, style);
  
  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.7,
        max_tokens: 1000
      })
    });
    
    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error?.message || `API error: ${response.status}`);
    }
    
    const data = await response.json();
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      throw new Error('Invalid response from OpenAI API');
    }
    
    return {
      success: true,
      rewrittenText: data.choices[0].message.content.trim()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// Call translate API
async function callTranslateAPI(text, targetLang, apiKey) {
  const prompt = `Translate the following text to ${targetLang}. 

CRITICAL RULES:
- DO NOT add quotation marks, brackets, or any punctuation that wasn't in the original
- DO NOT add subject lines, headers, titles, or any formatting that wasn't in the original text
- Preserve the original formatting, including line breaks and paragraph structure
- Maintain the same spacing and line breaks as the original
- Only return the translated text, nothing else
- Translate naturally and accurately
- Keep the text clean and well-formatted - remove excessive spacing or line breaks

IMPORTANT: The translated text should have the same formatting structure as the original. If the original has single line breaks, keep single line breaks. If it has double line breaks, keep double line breaks. But remove any excessive spacing (more than 2 consecutive line breaks).

Text: "${text}"`;

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: prompt }],
        temperature: 0.3,
        max_tokens: 1000
      })
    });
    
    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error?.message || `API error: ${response.status}`);
    }
    
    const data = await response.json();
    if (!data.choices || !data.choices[0] || !data.choices[0].message) {
      throw new Error('Invalid response from OpenAI API');
    }
    
    return {
      success: true,
      translatedText: data.choices[0].message.content.trim()
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// Get prompt for style
function getPromptForStyle(text, style) {
  if (style === 'professional' || style === 'improve') {
    return `Improve the wording and clarity of the following text while preserving ALL factual information exactly as written. 

ABSOLUTE REQUIREMENTS - NO EXCEPTIONS:
- DO NOT add quotation marks, double quotes, single quotes, or any quotation punctuation
- DO NOT add brackets, parentheses (unless they were in the original), or any extra punctuation
- DO NOT add subject lines, headers, titles, or any formatting that wasn't in the original text
- DO NOT change any numbers, percentages, or statistics (e.g., "4.35%", "4 months", "1/10 people")
- DO NOT change any dates, deadlines, or timeframes
- DO NOT change any conditions, terms, or legal language
- DO NOT change any specific facts, data, or figures
- Only return the improved text, nothing else - no quotes, no explanations, just the rewritten text

IMPORTANT: This is a selected portion of text from a larger message. Rewrite it naturally as if it's part of the original message. Do NOT wrap it in quotes or treat it as a quote. Just return the improved version of the text itself.

Only improve the grammar, sentence structure, flow, and word choice. Make it clearer and more professional while keeping all factual content identical.

CAPITALIZATION FIXES:
- Capitalize the first letter of the first word in each sentence
- Fix any capitalization errors throughout the text
- Ensure proper capitalization of proper nouns, names, and titles
- Maintain consistent capitalization

IMPORTANT: Preserve the original formatting, including line breaks and paragraph structure. If the original has blank lines or specific spacing, maintain that structure.

Text: "${text}"`;
  } else if (style === 'factcheck') {
    return `You are a professional fact-checker. Your job is to verify EVERY fact, statistic, and claim in the text below and CORRECT any inaccuracies with verified, accurate information.

ABSOLUTE REQUIREMENTS - NO EXCEPTIONS:
- DO NOT add quotation marks, double quotes, single quotes, or any quotation punctuation
- DO NOT add brackets, parentheses (unless they were in the original), or any extra punctuation
- DO NOT add subject lines, headers, titles, or any formatting that wasn't in the original text
- DO NOT add explanations about data availability, research limitations, or meta-commentary (e.g., "statistics are not readily available", "data suggests", "research indicates that data is limited")
- Only return the fact-checked and improved text, nothing else - no quotes, no explanations, just the rewritten text

IMPORTANT: This is a selected portion of text from a larger message. Rewrite it naturally as if it's part of the original message. Do NOT wrap it in quotes or treat it as a quote. Just return the corrected version of the text itself.

Write naturally as if the facts were always correct - don't explain your fact-checking process.

FACT-CHECKING REQUIREMENTS:
1. Identify ALL facts, statistics, percentages, numbers, and claims in the text
2. For EACH fact/statistic, determine if it's accurate:
   - If you know the correct statistic, REPLACE the incorrect one with the accurate data
   - If a claim is unverifiable or likely false, simply REMOVE it and rewrite the sentence naturally without that claim
   - Do NOT add explanations about why you removed something or what data is available
3. Rewrite the text with:
   - ALL incorrect facts/statistics CORRECTED with accurate information (when available)
   - Unverifiable claims REMOVED (without explanation)
   - Improved clarity, grammar, and professional tone
   - Write as if the corrected version is the original - no meta-commentary
   - Preserve the original formatting, including line breaks and paragraph structure
   - Fix capitalization errors (capitalize first word of sentences, proper nouns, etc.)

Text: "${text}"`;
  } else if (style === 'persuasive') {
    return `Rewrite the following text to be persuasive and sales-oriented, but keep it concise and professional. Make it compelling without being overly long or cheesy.

ABSOLUTE REQUIREMENTS - NO EXCEPTIONS:
- DO NOT add quotation marks, double quotes, single quotes, or any quotation punctuation
- DO NOT add brackets, parentheses (unless they were in the original), or any extra punctuation
- DO NOT add subject lines, headers, titles, or any formatting that wasn't in the original text
- Only return the persuasive version of the text, nothing else - no quotes, no explanations, just the rewritten text
- Preserve all factual information (numbers, percentages, dates, conditions) exactly as written
- Keep it CONCISE - similar length to the original, not much longer
- Maintain the same tone and structure as the original

IMPORTANT: This is a selected portion of text from a larger message. Rewrite it naturally as if it's part of the original message. Do NOT wrap it in quotes or treat it as a quote. Just return the persuasive version of the text itself.

PERSUASIVE WRITING GUIDELINES:
- Use confident, professional language
- Highlight benefits and value
- Create subtle urgency without being pushy
- Be persuasive but not overly salesy or cheesy
- Keep it natural and professional
- Don't add excessive fluff or lengthy explanations
- Fix capitalization errors (capitalize first word of sentences, proper nouns, etc.)

The goal is to make it more compelling for sales, but keep it concise and professional - not a long sales pitch.

IMPORTANT: Preserve the original formatting, including line breaks and paragraph structure.

Text: "${text}"`;
  }
  
  return `Rewrite the following text to improve clarity, grammar, and flow while maintaining the original meaning.\n\nText: "${text}"`;
}

// Set buttons enabled/disabled
function setButtonsEnabled(enabled) {
  document.getElementById('btn-professional').disabled = !enabled;
  document.getElementById('btn-factcheck').disabled = !enabled;
  document.getElementById('btn-persuasive').disabled = !enabled;
  document.getElementById('btn-translate').disabled = !enabled;
}

// Show status message
function showStatus(message, type) {
  const statusDiv = document.getElementById('status-message');
  statusDiv.textContent = message;
  statusDiv.className = 'status-message status-' + type;
  setTimeout(() => {
    statusDiv.textContent = '';
    statusDiv.className = 'status-message';
  }, 3000);
}

// Load settings
async function loadSettings() {
  // Load API key status
  const apiKey = await getApiKey();
  if (apiKey) {
    showStatus('API key configured', 'success');
  } else {
    showStatus('API key not set', 'warning');
  }
  
  // Load feature toggles
  try {
    const enabled = Office.context.roamingSettings.get('enabledFeatures');
    if (enabled) {
      document.getElementById('toggle-professional').checked = enabled.professional !== false;
      document.getElementById('toggle-factcheck').checked = enabled.factcheck !== false;
      document.getElementById('toggle-persuasive').checked = enabled.persuasive !== false;
      document.getElementById('toggle-translate').checked = enabled.translate !== false;
    }
  } catch (error) {
    console.error('Error loading feature toggles:', error);
  }
  updateButtonVisibility();
}

// Save feature toggles
function saveFeatureToggles() {
  const features = {
    professional: document.getElementById('toggle-professional').checked,
    factcheck: document.getElementById('toggle-factcheck').checked,
    persuasive: document.getElementById('toggle-persuasive').checked,
    translate: document.getElementById('toggle-translate').checked
  };
  
  Office.context.roamingSettings.set('enabledFeatures', features);
  Office.context.roamingSettings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      updateButtonVisibility();
      showStatus('Settings saved', 'success');
    }
  });
}

// Update button visibility based on toggles
function updateButtonVisibility() {
  try {
    const enabled = Office.context.roamingSettings.get('enabledFeatures') || 
      { professional: true, factcheck: true, persuasive: true, translate: true };
    
    document.getElementById('btn-professional').style.display = 
      enabled.professional !== false ? 'inline-block' : 'none';
    document.getElementById('btn-factcheck').style.display = 
      enabled.factcheck !== false ? 'inline-block' : 'none';
    document.getElementById('btn-persuasive').style.display = 
      enabled.persuasive !== false ? 'inline-block' : 'none';
    document.getElementById('btn-translate').style.display = 
      enabled.translate !== false ? 'inline-block' : 'none';
  } catch (error) {
    console.error('Error updating button visibility:', error);
    // Show all buttons by default
    document.getElementById('btn-professional').style.display = 'inline-block';
    document.getElementById('btn-factcheck').style.display = 'inline-block';
    document.getElementById('btn-persuasive').style.display = 'inline-block';
    document.getElementById('btn-translate').style.display = 'inline-block';
  }
}

// Open settings
function openSettings() {
  const apiKey = prompt('Enter your OpenAI API key:', '');
  if (apiKey && apiKey.trim()) {
    saveApiKey(apiKey.trim()).then((success) => {
      if (success) {
        showStatus('API key saved successfully!', 'success');
      } else {
        showStatus('Error saving API key.', 'error');
      }
    });
  }
}

