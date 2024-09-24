// Define interfaces for the JSON data structure
interface TagData {
  tag: string;
  data: string[];
}

// Store tags and datasets
let tagData: TagData[] = [];
let currentDatasetIndex = 0;
let jsonFilename = '';

// Plugin UI window size: 600 x 448
figma.showUI(__html__, { width: 600, height: 448 });

// Listen for messages from the UI
figma.ui.onmessage = async (msg) => {
  if (msg.type === 'ui-ready') {
    // Load saved data from client storage
    const savedTagData = await figma.clientStorage.getAsync('tagData');
    const savedFilename = await figma.clientStorage.getAsync('jsonFilename');
    const savedDatasetIndex = await figma.clientStorage.getAsync('currentDatasetIndex');

    if (savedTagData && Array.isArray(savedTagData)) {
      tagData = savedTagData;
      jsonFilename = savedFilename || '';
      currentDatasetIndex = savedDatasetIndex || 0;

      // Ensure the currentDatasetIndex is within valid range
      if (tagData.length > 0 && currentDatasetIndex >= tagData[0].data.length) {
        currentDatasetIndex = 0;
      }

      figma.ui.postMessage({
        type: 'update-dataset-index',
        index: currentDatasetIndex,
        total: tagData.length > 0 ? tagData[0].data.length : 0,
        tagData: tagData,
        filename: jsonFilename,
      });
    }
  } else if (msg.type === 'import-json') {
    tagData = msg.tagData;
    jsonFilename = msg.filename;
    currentDatasetIndex = 0;

    // Save the data to client storage
    await figma.clientStorage.setAsync('tagData', tagData);
    await figma.clientStorage.setAsync('jsonFilename', jsonFilename);
    await figma.clientStorage.setAsync('currentDatasetIndex', currentDatasetIndex);

    figma.ui.postMessage({
      type: 'update-dataset-index',
      index: currentDatasetIndex,
      total: tagData.length > 0 ? tagData[0].data.length : 0,
      tagData: tagData,
      filename: jsonFilename,
    });
  } else if (msg.type === 'export-json') {
    figma.ui.postMessage({
      type: 'save-json',
      tagData: tagData,
      filename: jsonFilename, // Include the filename
    });
  } else if (msg.type === 'unload-json') {
    // Clear the data
    tagData = [];
    jsonFilename = '';
    currentDatasetIndex = 0;
    // Remove data from client storage
    await figma.clientStorage.setAsync('tagData', []);
    await figma.clientStorage.setAsync('jsonFilename', '');
    await figma.clientStorage.setAsync('currentDatasetIndex', 0);

    // Update the UI
    figma.ui.postMessage({
      type: 'update-dataset-index',
      index: 0,
      total: 0,
      tagData: [],
      filename: '',
    });
  } else if (msg.type === 'next-dataset') {
    if (tagData.length > 0 && currentDatasetIndex < tagData[0].data.length - 1) {
      currentDatasetIndex++;
      // Save the current dataset index
      await figma.clientStorage.setAsync('currentDatasetIndex', currentDatasetIndex);

      figma.ui.postMessage({
        type: 'update-dataset-index',
        index: currentDatasetIndex,
        total: tagData[0].data.length,
        tagData: tagData,
        filename: jsonFilename,
      });
    }
  } else if (msg.type === 'prev-dataset') {
    if (currentDatasetIndex > 0) {
      currentDatasetIndex--;
      // Save the current dataset index
      await figma.clientStorage.setAsync('currentDatasetIndex', currentDatasetIndex);

      figma.ui.postMessage({
        type: 'update-dataset-index',
        index: currentDatasetIndex,
        total: tagData[0].data.length,
        tagData: tagData,
        filename: jsonFilename,
      });
    }
  } else if (msg.type === 'start') {
    const selection = figma.currentPage.selection;

    if (selection.length === 0) {
      figma.notify('Please select at least one layer.');
    } else if (tagData.length === 0) {
      figma.notify('Please import a JSON file with tags and datasets.');
    } else {
      // Create a mapping from tags to data for the current dataset
      const tagToData: { [key: string]: string } = {};
      for (const tagEntry of tagData) {
        const dataValue = tagEntry.data[currentDatasetIndex];
        if (dataValue && dataValue.trim() !== '') {
          tagToData[tagEntry.tag] = dataValue;
        }
      }

      // Recursive function to traverse and process nodes
      const traverse = async (node: SceneNode) => {
        // If the node has children, process them
        if ('children' in node) {
          for (const child of node.children) {
            await traverse(child);
          }
        }

        // If the node is a TextNode, replace tags
        if (node.type === 'TEXT') {
          const textNode = node as TextNode;

          // Load all fonts used in the text node
          await Promise.all(
            textNode
              .getRangeAllFontNames(0, textNode.characters.length)
              .map(figma.loadFontAsync)
          );

          // Perform the text replacement for each tag
          let newText = textNode.characters;
          for (const tag in tagToData) {
            if (Object.prototype.hasOwnProperty.call(tagToData, tag)) {
              const dataValue = tagToData[tag];
              // Escape special characters in the tag for regex
              const escapedTag = tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

              // Limit the replacement to avoid overloading (set limit, e.g., 1000 matches)
              const regex = new RegExp(escapedTag, 'g');

              const maxReplacements = 1000; // Safety limit on replacements
              let matchCount = (newText.match(regex) || []).length;

              if (matchCount <= maxReplacements) {
                newText = newText.replace(regex, dataValue);
              } else {
                console.warn(`Too many replacements for tag: ${tag}. Limiting to ${maxReplacements}.`);
                newText = newText.replace(regex, (match, offset) => {
                  if (--matchCount > 0) return dataValue;
                  return match; // Skip replacements once limit is hit
                });
              }
            }
          }
          textNode.characters = newText;
        }
      };

      // Process each selected node
      for (const node of selection) {
        await traverse(node);
      }

      figma.notify(`Processed dataset #${currentDatasetIndex + 1}.`);
    }
  }
};
