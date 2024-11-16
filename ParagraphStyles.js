#target indesign

// Function to get paragraph styles and font sizes across all pages (skipping the first 6 pages)
function getParagraphStylesInDocument() {
    var doc = app.activeDocument;
    
    // Create a new text file to store the results
    var file = new File(Folder.desktop + "/ParagraphStylesErrors.txt");
    file.encoding = "UTF-8";
    file.open("w"); // Open file for writing

    var foundErrors = false; // To track if any errors were found

    // Iterate through all pages in the document, starting from page 7 (index 6, as pages are zero-indexed)
    for (var p = 6; p < doc.pages.length; p++) {
        var targetPage = doc.pages[p];
        var pageNumber = targetPage.name; // Get the actual document page number

        // Iterate through all text frames on the page
        for (var i = 0; i < targetPage.textFrames.length; i++) {
            var textFrame = targetPage.textFrames[i];
            var paragraphs = textFrame.paragraphs;

            // Iterate through all paragraphs in the text frame
            for (var j = 0; j < paragraphs.length; j++) {
                var paragraph = paragraphs[j];
                var paragraphStyle = paragraph.appliedParagraphStyle.name;
                var styleFontSize = paragraph.appliedParagraphStyle.pointSize; // Get the font size defined in the paragraph style
                
                // Get the first word of the paragraph for better identification
                var firstWord = paragraph.words.length > 0 ? paragraph.words[0].contents : "N/A";

                // Variables to track ranges of characters with different font sizes
                var startRange = null;
                var endRange = null;
                var previousFontSize = null;

                // Iterate through characters and check their font size
                for (var k = 0; k < paragraph.characters.length; k++) {
                    var characterFontSize = paragraph.characters[k].pointSize;

                    if (characterFontSize != styleFontSize) {
                        // Start or continue a range if font size differs from paragraph style
                        if (startRange === null) {
                            startRange = k; // Mark the start of the range
                        }
                        endRange = k; // Update the end of the range
                        previousFontSize = characterFontSize;
                    } else {
                        // If the current character matches the style font size, check if we were tracking a different range
                        if (startRange !== null) {
                            // Build the differing content
                            var differingContent = paragraph.contents.substring(startRange, endRange + 1);

                            // Log the range of different font size and include the page number
                            file.writeln("Page " + pageNumber + ". In paragraph starting with '" + firstWord + "', content: '" + differingContent + "' has a different font size: " + previousFontSize + " pt");
                            
                            foundErrors = true; // Mark that an error was found

                            // Reset range tracking
                            startRange = null;
                            endRange = null;
                        }
                    }
                }

                // After the loop, if there's still a range being tracked, log it
                if (startRange !== null) {
                    var differingContent = paragraph.contents.substring(startRange, endRange + 1);
                    file.writeln("Page " + pageNumber + ". In paragraph starting with '" + firstWord + "', Content: '" + paragraph.contents + "' has a different font size: " + previousFontSize + " pt for the text: '" + differingContent + "'");
                    foundErrors = true;
                }
            }
        }
    }

    file.close(); // Close the file after writing

    // Alert if no errors were found
    if (!foundErrors) {
        alert("No errors found in paragraph styles.");
    } else {
        alert("Paragraph styles with differing font sizes have been logged.");
    }
}

// Run the function for the whole document, skipping the first 6 pages
getParagraphStylesInDocument();
