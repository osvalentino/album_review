// Google Apps Script for my Google Slides
// Its purpose is to export text data from slide into a Google Sheet
// in the Google Sheet, each row represents a slide
// each column represents a text box which are consistent

function exportToSheet() {
  // Getting Album Review Slides
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();
  
  // getting Album Reviews Data Sheet
  const sheetUrl = 'https://docs.google.com/spreadsheets/d/17I3OvnBi4OnVEHwU6rUbvNAtlaSK3j25RCozVECPVAU/edit?gid=0#gid=0';
  let sheet;
  sheet = SpreadsheetApp.openByUrl(sheetUrl).getActiveSheet();
  
  // clearing it to refresh everything
  sheet.clear();
  sheet.appendRow(["album", "artist_year", "length", "genre", "country_lang", "rating", "all_song_ratings", "all_songs", "caption", "fav_song", "recommend", "superlative1", "superlative2"]);

  // iterating per slide
  for (let i = 0; i < slides.length; i++) {
    const shapes = slides[i].getShapes();

    // sort shapes left-to-right, then top-to-bottom
    const sorted = shapes.sort((a, b) => {
      const leftDiff = a.getLeft() - b.getLeft();
      if (Math.abs(leftDiff) > 5) { // allow small tolerance
        return leftDiff;
      }
      return a.getTop() - b.getTop();
    });

    // iterating per text box
    const row = [];
    for (let shape of sorted) {
      if (shape.getText()) {
        const text = shape.getText().asString().trim();
        if (text) row.push(text);
      }
    }
    sheet.appendRow(row);
  }
  Logger.log("Export complete");
}

