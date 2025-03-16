function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Search', 'showSearchSidebar')
      .addToUi();
}

// Show a sidebar with search functionality
function showSearchSidebar() {
  var htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 0;
        font-size: 12px;
      }
      .container {
        padding: 8px;
      }
      .header {
        background: #f1f3f4;
        padding: 12px;
        border-bottom: 1px solid #dadce0;
        font-size: 14px;
        font-weight: 500;
      }
      .search-container {
        padding: 12px;
        border-bottom: 1px solid #dadce0;
      }
      .search-box {
        display: flex;
        width: 100%;
        margin-bottom: 5px;
      }
      .search-input {
        flex-grow: 1;
        padding: 8px;
        border: 1px solid #ddd;
        border-radius: 4px 0 0 4px;
        font-size: 12px;
        outline: none;
      }
      .search-input:focus {
        border-color: #4285f4;
        box-shadow: 0 0 0 1px #4285f4;
      }
      .search-button {
        background: #4285f4;
        color: white;
        border: none;
        border-radius: 0 4px 4px 0;
        padding: 8px 12px;
        font-size: 12px;
        cursor: pointer;
      }
      .search-button:hover {
        background: #3b78e7;
      }
      .hint {
        font-size: 11px;
        color: #666;
        margin-top: 5px;
      }
      .results-header {
        padding: 8px 12px;
        font-size: 13px;
        color: #5f6368;
        border-bottom: 1px solid #dadce0;
      }
      .result-count {
        font-weight: bold;
      }
      .no-results {
        padding: 20px;
        text-align: center;
        color: #5f6368;
      }
      .result-item {
        padding: 10px 12px;
        border-bottom: 1px solid #dadce0;
        cursor: pointer;
      }
      .result-item:hover {
        background-color: #f8f9fa;
      }
      .result-text {
        margin-bottom: 4px;
      }
      .result-tags {
        color: #188038;
        font-size: 11px;
      }
      .result-location {
        color: #5f6368;
        font-size: 11px;
      }
      #results-container {
        overflow-y: auto;
        max-height: calc(100vh - 200px);
      }
    </style>
    
    <div class="container">
      <div class="header">Tag Search</div>
      
      <div class="search-container">
        <div class="search-box">
          <input type="text" id="searchTerm" class="search-input" placeholder="Search tags...">
          <button onclick="search()" class="search-button">Search</button>
        </div>
        <div class="hint">
          Examples: &lt;tag1&gt;|&lt;tag2&gt; (OR), &lt;tag1&gt;&amp;&lt;tag2&gt; (AND)
        </div>
      </div>
      
      <div id="results-header" class="results-header" style="display: none;">
        Search: <span id="search-term-display"></span>
        <div id="result-count" class="result-count"></div>
      </div>
      
      <div id="results-container"></div>
    </div>
    
    <script>
      // Focus on search input when sidebar opens
      document.addEventListener('DOMContentLoaded', function() {
        document.getElementById("searchTerm").focus();
      });
      
      // Allow pressing Enter to search
      document.getElementById("searchTerm").addEventListener("keypress", function(event) {
        if (event.key === "Enter") {
          search();
        }
      });
      
      function search() {
        var searchTerm = document.getElementById("searchTerm").value;
        if (!searchTerm) return;
        
        // Show loading indicator
        document.getElementById("results-container").innerHTML = '<div style="text-align: center; padding: 20px;">Searching...</div>';
        document.getElementById("results-header").style.display = "block";
        document.getElementById("search-term-display").textContent = searchTerm;
        
        google.script.run
          .withSuccessHandler(displayResults)
          .searchTags(searchTerm);
      }
      
      function displayResults(results) {
        var resultsContainer = document.getElementById("results-container");
        document.getElementById("result-count").textContent = results.length + " results found";
        
        if (results.length === 0) {
          resultsContainer.innerHTML = '<div class="no-results">No matching tags found</div>';
          return;
        }
        
        var html = '';
        for (var i = 0; i < results.length; i++) {
          var result = results[i];
          html += '<div class="result-item" onclick="navigateToResult(' + result.index + ')">' +
                 '<div class="result-text">' + escapeHtml(result.text) + '</div>' +
                 '<div class="result-location">Paragraph ' + result.paragraph + '</div>' +
                 '</div>';
        }
        
        resultsContainer.innerHTML = html;
      }
      
      function escapeHtml(unsafe) {
        return unsafe
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;")
          .replace(/'/g, "&#039;");
      }
      
      function navigateToResult(paragraphIndex) {
        google.script.run.navigateToParagraph(paragraphIndex);
      }
    </script>
  `);
  
  htmlOutput.setTitle('Tag Search');
  DocumentApp.getUi().showSidebar(htmlOutput);
}

function searchTags(searchTerm) {
  // Get the document content
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  var results = [];
  
  // Search through all paragraphs
  for (var i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();
    
    // Look for "tags:" in the paragraph
    var tagsIndex = text.toLowerCase().indexOf("tags:");
    if (tagsIndex !== -1) {
      // Extract the text after "tags:" (including the colon)
      var tagsText = text.substring(tagsIndex + 5).trim(); // +5 to skip "tags:"
      
      // Search for the term in the tags text
      if (tagsText.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1) {
        results.push({
          text: text.substring(0, 100) + (text.length > 100 ? "..." : ""),
          paragraph: i + 1,
          index: i
        });
      }
    }
  }

  return results;
}

function navigateToParagraph(paragraphIndex) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  if (paragraphIndex < paragraphs.length) {
    var paragraph = paragraphs[paragraphIndex];
    var position = doc.newPosition(paragraph, 0);
    doc.setCursor(position);
  }
}
