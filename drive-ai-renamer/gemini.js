/**
 * Uploads a file to Gemini File API.
 * @param {Blob} fileBlob The file blob to upload.
 * @param {string} filename The filename.
 * @returns {string | null} The file URI, or null on failure.
 */
function uploadFileToGemini(fileBlob, filename) {
  const apiKey = getApiKey();
  const url = 'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + apiKey;
  
  // Gemini File API requires multipart/form-data
  const boundary = '----WebKitFormBoundary' + Utilities.getUuid();
  
  // Build metadata JSON
  const metadata = {
    file: {
      displayName: filename
    }
  };
  const metadataJson = JSON.stringify(metadata);
  
  // Build multipart payload as bytes
  const fileBytes = fileBlob.getBytes();
  const CRLF = '\r\n';
  const boundaryLine = '--' + boundary + CRLF;
  
  // Build the multipart payload parts
  const metadataHeader = boundaryLine +
    'Content-Disposition: form-data; name="metadata"' + CRLF +
    'Content-Type: application/json' + CRLF + CRLF;
  const metadataHeaderBytes = Utilities.newBlob(metadataHeader).getBytes();
  
  const metadataJsonBytes = Utilities.newBlob(metadataJson).getBytes();
  
  const fileSeparator = CRLF + boundaryLine;
  const fileSeparatorBytes = Utilities.newBlob(fileSeparator).getBytes();
  
  const fileHeader = 'Content-Disposition: form-data; name="file"; filename="' + filename + '"' + CRLF +
    'Content-Type: application/pdf' + CRLF + CRLF;
  const fileHeaderBytes = Utilities.newBlob(fileHeader).getBytes();
  
  const closingBoundary = CRLF + '--' + boundary + '--' + CRLF;
  const closingBoundaryBytes = Utilities.newBlob(closingBoundary).getBytes();
  
  // Combine all parts into a single byte array
  // In Apps Script, getBytes() returns a byte array that we need to convert properly
  let payloadBytes = [];
  
  // Helper function to convert byte array to regular array
  function toArray(bytes) {
    const arr = [];
    for (let i = 0; i < bytes.length; i++) {
      arr.push(bytes[i]);
    }
    return arr;
  }
  
  payloadBytes = payloadBytes.concat(toArray(metadataHeaderBytes));
  payloadBytes = payloadBytes.concat(toArray(metadataJsonBytes));
  payloadBytes = payloadBytes.concat(toArray(fileSeparatorBytes));
  payloadBytes = payloadBytes.concat(toArray(fileHeaderBytes));
  payloadBytes = payloadBytes.concat(toArray(fileBytes));
  payloadBytes = payloadBytes.concat(toArray(closingBoundaryBytes));
  
  const payloadBlob = Utilities.newBlob(payloadBytes);
  
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'multipart/form-data; boundary=' + boundary
    },
    'payload': payloadBlob.getBytes(),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      // Return the file URI (file.name)
      return json.file.name;
    } else {
      Logger.log(`Gemini File Upload Error ${responseCode}: ${responseText}`);
      return null;
    }
  } catch (e) {
    Logger.log(`Exception during file upload: ${e}`);
    return null;
  }
}

/**
 * Calls the Gemini API with a structured prompt.
 * @param {string} userPrompt The user's instruction.
 * @param {string[]} filenames The list of old filenames.
 * @param {Array<{filename: string, mimeType: string, base64Data: string}>} fileData Optional array of file data with base64-encoded PDFs/images for content analysis.
 * @returns {string[] | null} An array of new filenames, or null on failure.
 */
function callGemini(userPrompt, filenames, fileData = null) {
  const apiKey = getApiKey();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';

  // Build the prompt based on whether we have file content
  let systemPrompt;
  let parts = [];
  
  if (fileData && fileData.length > 0) {
    // We have file content - instruct Gemini to analyze the files
    systemPrompt = `
      You are a file renaming assistant. You will be given a user's rule, a JSON list of filenames, and the actual file contents (PDFs and/or images).
      Your job is to analyze the file contents and apply the rule to EVERY filename, returning ONLY a valid JSON array of the new names, in the exact same order.
      
      - Analyze the content of each file (PDFs and images) to extract relevant information for renaming.
      - Apply the user's rule based on the file content.
      - Do not add any commentary.
      - Do not add markdown (like \`\`\`json).
      - Ensure new filenames are valid (no '/' or '\' characters).
      - If a file should not be renamed, return its original name in the list.
      - The output MUST be a JSON array of strings.

      User Rule: "${userPrompt}"
      
      Files to rename (in order): ${JSON.stringify(filenames)}
      
      Your Response:
    `;
    
    // Add all files inline first, then the prompt
    // According to Gemini docs: https://ai.google.dev/gemini-api/docs/document-processing#prompt-multiple
    // and https://ai.google.dev/gemini-api/docs/image-understanding#supported-formats
    fileData.forEach((file) => {
      if (file.base64Data && file.mimeType) {
        parts.push({
          "inlineData": {
            "mimeType": file.mimeType,
            "data": file.base64Data
          }
        });
      }
    });
    
    // Add the text prompt after all files
    parts.push({"text": systemPrompt});
  } else {
    // No file content - use original behavior
    systemPrompt = `
      You are a file renaming assistant. You will be given a user's rule and a JSON list of filenames.
      Your job is to apply the rule to EVERY filename and return ONLY a valid JSON array of the new names, in the exact same order.
      
      - Do not add any commentary.
      - Do not add markdown (like \`\`\`json).
      - Ensure new filenames are valid (no '/' or '\' characters).
      - If a file should not be renamed, return its original name in the list.
      - The output MUST be a JSON array of strings.

      User Rule: "${userPrompt}"
      
      Files: ${JSON.stringify(filenames)}
      
      Your Response:
    `;
    parts.push({"text": systemPrompt});
  }

  const payload = {
    "contents": [{
      "parts": parts
    }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "temperature": 0.0, // We want deterministic, not creative, responses
    }
  };

  const options = {
    'method': 'post',
    'headers': {
      'x-goog-api-key': apiKey
    },
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true // We'll handle errors ourselves
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      // The Gemini API returns a complex object. We need to parse our *expected* JSON from its content.
      // The actual response content is inside: response.candidates[0].content.parts[0].text
      const textOutput = json.candidates[0].content.parts[0].text;
      return JSON.parse(textOutput); // This should be our array: ["new1.pdf", "new2.doc"]
    } else {
      Logger.log(`Gemini Error ${responseCode}: ${responseText}`);
      return null;
    }
  } catch (e) {
    Logger.log(`Exception during API call: ${e}`);
    return null;
  }
}