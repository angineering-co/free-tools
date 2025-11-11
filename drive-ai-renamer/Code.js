// --- HELPER ---
function getApiKey() {
  return PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
}

/**
 * Main entry point. Called when items are selected in Drive.
 * @param {Object} e The event object.
 * @returns {Card[]} An array of cards to display.
 */
function onDriveItemsSelected(e) {
  const selectedItems = e.drive.selectedItems;

  // Case 1: Single folder selected
  if (
    selectedItems.length === 1 &&
    selectedItems[0].mimeType === "application/vnd.google-apps.folder"
  ) {
    const selectedFolderId = selectedItems[0].id;
    return [buildMainRenamingCard(selectedFolderId, null)];
  }

  // Case 2: Multiple files selected (must be in the same folder)
  if (selectedItems.length > 0) {
    // Check if all selected items are files (not folders)
    const allAreFiles = selectedItems.every(
      (item) => item.mimeType !== "application/vnd.google-apps.folder"
    );

    if (allAreFiles) {
      // Get parent folders for all files and verify they're all in the same folder
      const parentFolderIds = new Set();
      const selectedFileIds = [];

      for (let i = 0; i < selectedItems.length; i++) {
        const item = selectedItems[i];
        try {
          const file = DriveApp.getFileById(item.id);
          const parentFolders = file.getParents();

          if (!parentFolders.hasNext()) {
            return [
              createErrorCard(
                `File "${item.title}" is not in a folder. Please select files from a folder.`
              ),
            ];
          }

          const parentFolder = parentFolders.next();
          parentFolderIds.add(parentFolder.getId());
          selectedFileIds.push(item.id);
        } catch (err) {
          Logger.log(`Error processing file ${item.id}: ${err}`);
          return [
            createErrorCard(
              `Error processing selected files. Please try again.`
            ),
          ];
        }
      }

      // Verify all files are in the same folder
      if (parentFolderIds.size !== 1) {
        return [
          createErrorCard(
            "Please select files from the same folder, or select a single folder."
          ),
        ];
      }

      const folderId = Array.from(parentFolderIds)[0];
      return [buildMainRenamingCard(folderId, selectedFileIds)];
    }
  }

  // Invalid selection
  return [
    createErrorCard(
      "Please select a single folder, or multiple files from the same folder."
    ),
  ];
}

/**
 * Builds the main UI card (Card 1)
 * @param {string} folderId The ID of the folder containing the files.
 * @param {string[]|null} selectedFileIds Optional array of selected file IDs. If null, processes all files in folder.
 * @returns {Card}
 */
function buildMainRenamingCard(folderId, selectedFileIds) {

  // TODO: Add an alert to stop the flow if more than 20 files are selected

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Gemini File Renamer"));

  // API Key Section
  const apiKeySection = CardService.newCardSection();
  apiKeySection.setHeader("API Configuration");

  const hasApiKey = !!getApiKey();
  if (hasApiKey) {
    apiKeySection.addWidget(
      CardService.newTextParagraph().setText(
        '<font color="#34a853">✓ API Key is configured</font>'
      )
    );
  } else {
    apiKeySection.addWidget(
      CardService.newTextParagraph().setText(
        '<font color="#ea4335">⚠ API Key required</font>'
      )
    );
  }

  apiKeySection.addWidget(
    CardService.newTextInput()
      .setFieldName("api_key")
      .setTitle("Gemini API Key")
      .setHint("Enter your Gemini API key")
      .setValue("")
  ); // Clear the field

  const saveApiKeyActionParams = { folderId: folderId };
  if (selectedFileIds) {
    saveApiKeyActionParams.selectedFileIds = JSON.stringify(selectedFileIds);
  }
  const saveApiKeyAction = CardService.newAction()
    .setFunctionName("handleSaveApiKey")
    .setParameters(saveApiKeyActionParams);

  apiKeySection.addWidget(
    CardService.newButtonSet().addButton(
      CardService.newTextButton()
        .setText("Save API Key")
        .setOnClickAction(saveApiKeyAction)
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    )
  );

  card.addSection(apiKeySection);

  // Renaming Section
  const section = CardService.newCardSection();
  section.setHeader("Rename Files");
  section.addWidget(
    // TODO: Add button to save the prompt to the user properties so we can use it later
    CardService.newTextInput()
      .setFieldName("rename_prompt")
      .setTitle("Renaming Prompt")
      .setHint('e.g., "Add Project-X- prefix"')
  );

  // Pass the folderId and selectedFileIds to the action handler
  const actionParams = { folderId: folderId };
  if (selectedFileIds) {
    actionParams.selectedFileIds = JSON.stringify(selectedFileIds);
  }
  const action = CardService.newAction()
    .setFunctionName("handlePreview")
    .setParameters(actionParams);

  section.addWidget(
    CardService.newButtonSet().addButton(
      CardService.newTextButton()
        .setText("Preview Renames")
        .setOnClickAction(action)
        .setDisabled(false)
    )
  );

  card.addSection(section);
  return card.build();
}

/**
 * Builds a simple card to show an error message.
 * @param {string} text The error to display.
 * @returns {Card}
 */
function createErrorCard(text) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Error"))
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(text)
      )
    )
    .build();
}

/**
 * Called when the "Save API Key" button is clicked.
 * @param {Object} e The event object.
 * @returns {ActionResponse}
 */
function handleSaveApiKey(e) {
  const apiKey = e.formInput.api_key;
  const folderId = e.parameters.folderId;
  const selectedFileIdsParam = e.parameters.selectedFileIds;
  const selectedFileIds = selectedFileIdsParam
    ? JSON.parse(selectedFileIdsParam)
    : null;

  if (!apiKey || apiKey.trim() === "") {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText("Please enter an API key.")
      )
      .build();
  }

  // Store the API key
  PropertiesService.getUserProperties().setProperty(
    "GEMINI_API_KEY",
    apiKey.trim()
  );

  // Rebuild the card to show updated status and clear the input field
  // Preserve the file selection if files were originally selected
  const updatedCard = buildMainRenamingCard(folderId, selectedFileIds);

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText("API key saved successfully!")
    )
    .setNavigation(CardService.newNavigation().updateCard(updatedCard))
    .build();
}

/**
 * Called when the "Preview Renames" button is clicked.
 * @param {Object} e The event object.
 * @returns {ActionResponse}
 */
function handlePreview(e) {
  const prompt = e.formInput.rename_prompt;
  const folderId = e.parameters.folderId;
  const selectedFileIdsParam = e.parameters.selectedFileIds;
  const selectedFileIds = selectedFileIdsParam
    ? JSON.parse(selectedFileIdsParam)
    : null;

  let filesToProcess = [];

  if (selectedFileIds) {
    // Process only selected files
    for (let i = 0; i < selectedFileIds.length; i++) {
      try {
        const file = DriveApp.getFileById(selectedFileIds[i]);
        filesToProcess.push(file);
      } catch (err) {
        Logger.log(`Failed to get file ${selectedFileIds[i]}: ${err}`);
      }
    }
  } else {
    // Process all files in folder
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    while (files.hasNext()) {
      filesToProcess.push(files.next());
    }
  }

  let oldFilenames = [];
  let fileIds = []; // We need to save the IDs for the "execute" step
  let fileData = []; // Store file data for Gemini (PDFs and images) - base64 encoded
  let skippedFiles = []; // Track skipped unsupported files
  const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20MB per file limit
  const MAX_TOTAL_SIZE = 20 * 1024 * 1024; // 20MB total limit for all files in request
  let totalEncodedSize = 0; // Track cumulative base64-encoded size

  // Supported file types for inline upload
  const SUPPORTED_MIME_TYPES = new Set([
    "application/pdf",
    "image/png",
    "image/jpeg",
    "image/webp",
    "image/heic",
    "image/heif",
  ]);

  // Process files: filter supported types and encode them as base64 for inline upload
  for (let i = 0; i < filesToProcess.length; i++) {
    const file = filesToProcess[i];
    const mimeType = file.getMimeType();
    const filename = file.getName();

    // Only process supported file types (PDFs and images)
    if (!SUPPORTED_MIME_TYPES.has(mimeType)) {
      skippedFiles.push(filename);
      continue;
    }

    // Check per-file size limit
    const fileSize = file.getSize();
    if (fileSize > MAX_FILE_SIZE) {
      skippedFiles.push(filename + " (too large)");
      continue;
    }

    // Read file and encode as base64 for inline upload
    try {
      const fileBlob = file.getBlob();
      const base64Data = Utilities.base64Encode(fileBlob.getBytes());

      // Check total size limit (base64 encoding increases size by ~33%)
      // We check the actual encoded size since that's what's sent in the request
      const encodedSize = base64Data.length;
      if (totalEncodedSize + encodedSize > MAX_TOTAL_SIZE) {
        skippedFiles.push(filename + " (total size limit exceeded)");
        continue;
      }

      totalEncodedSize += encodedSize;
      oldFilenames.push(filename);
      fileIds.push(file.getId());
      fileData.push({
        filename: filename,
        mimeType: mimeType,
        base64Data: base64Data,
      });
    } catch (err) {
      Logger.log(`Failed to process file ${filename}: ${err}`);
      skippedFiles.push(filename + " (processing failed)");
    }
  }

  if (oldFilenames.length === 0) {
    let message = selectedFileIds
      ? "No supported files (PDFs or images) found in the selection."
      : "No supported files (PDFs or images) found in the folder.";
    if (skippedFiles.length > 0) {
      message += " Skipped: " + skippedFiles.slice(0, 5).join(", ");
      if (skippedFiles.length > 5) {
        message += " and " + (skippedFiles.length - 5) + " more.";
      }
    }
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(message))
      .build();
  }

  // Call the Gemini API with file content
  const result = callGemini(prompt, oldFilenames, fileData);

  // --- Check for errors ---
  if (result && result.error) {
    // API returned an error object
    const errorMessage = result.details
      ? `${result.error}: ${result.details}`
      : result.error;
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(errorMessage))
      .build();
  }

  // --- Validation ---
  if (
    !result ||
    !Array.isArray(result) ||
    result.length !== oldFilenames.length
  ) {
    const errorMsg = !result
      ? "No response from AI. Please try again."
      : !Array.isArray(result)
      ? `Expected an array but got: ${typeof result}`
      : `Expected ${oldFilenames.length} filenames but got ${result.length}. Please try a clearer prompt.`;

    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(`Error: ${errorMsg}`)
      )
      .build();
  }

  const newFilenames = result;

  // --- Store the mapping for the *next* step ---
  // We use CacheService because it's temporary
  const cache = CacheService.getUserCache();
  const renameMap = {
    ids: fileIds,
    newNames: newFilenames,
  };
  cache.put("renameMap", JSON.stringify(renameMap), 300); // Store for 5 mins

  // Store skipped files info for display
  if (skippedFiles.length > 0) {
    cache.put("skippedFiles", JSON.stringify(skippedFiles), 300);
  }

  // Build the Preview Card (Card 2)
  const previewCard = buildPreviewCard(
    oldFilenames,
    newFilenames,
    skippedFiles
  );

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(previewCard))
    .build();
}

/**
 * Builds the Preview Card (Card 2)
 * @param {string[]} oldNames List of old filenames
 * @param {string[]} newNames List of new filenames
 * @param {string[]} skippedFiles Optional list of skipped files
 * @returns {Card}
 */
function buildPreviewCard(oldNames, newNames, skippedFiles = []) {
  const builder = CardService.newCardBuilder().setHeader(
    CardService.newCardHeader().setTitle("Confirm Renames")
  );

  const section = CardService.newCardSection().setHeader("Preview:");

  // Show skipped files info if any
  if (skippedFiles.length > 0) {
    const skippedText =
      skippedFiles.length === 1
        ? `Skipped: ${skippedFiles[0]}`
        : `Skipped ${skippedFiles.length} unsupported files or files that couldn't be processed.`;
    section.addWidget(
      CardService.newTextParagraph().setText(`<i>${skippedText}</i>`)
    );
    section.addWidget(CardService.newDivider());
  }

  // Create a (crude) table
  for (let i = 0; i < oldNames.length; i++) {
    // Show a truncated view for long lists
    if (i > 20) {
      section.addWidget(
        CardService.newTextParagraph().setText(
          `...and ${oldNames.length - i} more.`
        )
      );
      break;
    }
    section.addWidget(
      CardService.newTextParagraph().setText(
        `<b>${oldNames[i]}</b>  ->  <b>${newNames[i]}</b>`
      )
    );
  }

  // Add the "Confirm" and "Cancel" buttons
  const action = CardService.newAction().setFunctionName("handleExecuteRename");

  const buttonSet = CardService.newButtonSet()
    .addButton(
      CardService.newTextButton()
        .setText("Confirm & Rename")
        .setOnClickAction(action)
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    )
    .addButton(
      CardService.newTextButton()
        .setText("Cancel")
        .setOnClickAction(
          CardService.newAction().setFunctionName("onDriveItemsSelected")
        ) // Just reloads the main card
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
    );

  builder.addSection(section);
  builder.addSection(CardService.newCardSection().addWidget(buttonSet));

  return builder.build();
}

/**
 * Called when the "Confirm & Rename" button is clicked.
 * @param {Object} e The event object.
 * @returns {ActionResponse}
 */
function handleExecuteRename(e) {
  const cache = CacheService.getUserCache();
  const renameMapString = cache.get("renameMap");

  if (!renameMapString) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          "Error: Session timed out. Please preview again."
        )
      )
      .build();
  }

  const renameMap = JSON.parse(renameMapString);
  const fileIds = renameMap.ids;
  const newNames = renameMap.newNames;

  let successCount = 0;
  let errorCount = 0;

  // Rename files
  for (let i = 0; i < fileIds.length; i++) {
    try {
      const file = DriveApp.getFileById(fileIds[i]);
      file.setName(newNames[i]);
      successCount++;
    } catch (err) {
      Logger.log(`Failed to rename file ${fileIds[i]}: ${err}`);
      errorCount++;
    }
  }

  // Build the final Success Card (Card 3)
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Rename Complete"))
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(
          `<b>Success:</b> ${successCount} files renamed.<br><b>Failed:</b> ${errorCount} files.`
        )
      )
    )
    .build();

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}
