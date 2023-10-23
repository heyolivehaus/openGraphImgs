function fetchOgImagesAndWriteToSpecificColumn() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenGraph');
    // Read domains from column D (4th column) and checkboxes from column G (7th column)
    const domains = sheet.getRange(2, 4, 100, 1).getValues().flat();
    const checkboxes = sheet.getRange(2, 7, 100, 1).getValues().flat();
    
    for (let i = 0; i < domains.length; i++) {
      const domain = domains[i];
      const isChecked = checkboxes[i];
      
      if (domain && isChecked) { 
        try {
          const htmlContent = UrlFetchApp.fetch(domain).getContentText();
          const ogImageUrl = extractOgImage(htmlContent);
          
          if (ogImageUrl) {
            sheet.getRange(i + 2, 5).setValue(ogImageUrl);
          } else {
            console.warn(`No OG image found for domain: ${domain}`);
            sheet.getRange(i + 2, 6).setValue('No OG image found'); // column 6 will be where errors are printed out.
          }
        } catch (innerError) {
          console.error(`Error processing domain ${domain}: ${innerError.toString()}`);
          sheet.getRange(i + 2, 6).setValue(`Error: ${innerError.toString()}`); // column 6 will be where errors are printed out.
        }
      }
    }
  } catch (outerError) {
    console.error(`Unexpected error occurred: ${outerError.toString()}`);
    sheet.getRange(1, 6).setValue(`Unexpected error: ${outerError.toString()}`);
  }
}

function extractOgImage(htmlContent) {
  const ogImageMetaTagRegex = /<meta\s+property="og:image"\s+content="([^"]+)"/i;
  const match = htmlContent.match(ogImageMetaTagRegex);
  if (match && match[1]) {
    return match[1];
  }
  return null;
}
