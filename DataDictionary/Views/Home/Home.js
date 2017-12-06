(function () {
	"use strict";

	// The initialize function is run each time the page is loaded.
	Office.initialize = function (reason) {
		$(document).ready(function () {

			// Use this to check whether the API is supported in the Word client.
			if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
				// Do something that is only available via the new APIs
				$('#emerson').click(insertEmersonQuoteAtSelection);
				$('#checkhov').click(insertChekhovQuoteAtTheBeginning);
				$('#proverb').click(insertChineseProverbAtTheEnd);
				$('#insertTable').click(insertTable);
				$('#supportedVersion').html('This code is using Word 2016 or greater.');

				Word.run(function (context) {
					// Create a proxy object for the document body.
					var body = context.document.body;

					/*INSERTING A TABLE*/

					//body.insertTable(2, 2, 'End'); --doesn't work...'

					//body.insertHtml(
					//	"<table><tr><td>Column</td><td>Description</td></tr><tr><td>1</td><td>2</td></tr></table>",
					//	Word.InsertLocation.end
					//);

					/*SERACH FOR MATCHING TABLE NAMES, REPLACE WITH ACTUAL TABLE*/

					var searchResults = context.document.body.search('[DEBTOR4]', { matchCase: true });

					context.load(searchResults, 'text, font');

					// Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
					return context.sync().then(function () {

						for (var i = 0; i < searchResults.items.length; i++) {
							body.insertHtml(
								"<table><tr><td>Column</td><td>Description</td></tr><tr><td>1</td><td>2</td></tr></table>",
								Word.InsertLocation.replace
							);
						}

					});
				});
			}
			else {
				// Just letting you know that this code will not work with your version of Word.
				$('#supportedVersion').html('This code requires Word 2016 or greater.');
			}
		});
	};

	function insertTable() {
		Word.run(function (context) {
			// Create a proxy object for the document body.
			var body = context.document.body;

			//body.insertTable(2, 2, Word.InsertLocation.end, [["a", "b"], ["c", "d"]]);

			body.insertHtml(
				"<table><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></table>",
				Word.InsertLocation.end
			);

			// Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
			return context.sync();
		}).catch(function (e) {

			console.log(e.message);
		});
	}


	function insertEmersonQuoteAtSelection() {
		Word.run(function (context) {

			// Create a proxy object for the document.
			var thisDocument = context.document;

			// Queue a command to get the current selection.
			// Create a proxy range object for the selection.
			var range = thisDocument.getSelection();

			// Queue a command to replace the selected text.
			range.insertText('"Hitch your wagon to a star Chung2."\n', Word.InsertLocation.replace);

			// Synchronize the document state by executing the queued commands,
			// and return a promise to indicate task completion.
			return context.sync().then(function () {
				console.log('Added a quote from Ralph Waldo ');
			});
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}

	function insertChekhovQuoteAtTheBeginning() {
		Word.run(function (context) {

			// Create a proxy object for the document body.
			var body = context.document.body;

			// Queue a command to insert text at the start of the document body.
			body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

			// Synchronize the document state by executing the queued commands,
			// and return a promise to indicate task completion.
			return context.sync().then(function () {
				console.log('Added a quote from Anton Chekhov.');
			});
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}

	function insertChineseProverbAtTheEnd() {
		Word.run(function (context) {

			// Create a proxy object for the document body.
			var body = context.document.body;

			// Queue a command to insert text at the end of the document body.
			body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

			// Synchronize the document state by executing the queued commands,
			// and return a promise to indicate task completion.
			return context.sync().then(function () {
				console.log('Added a quote from a Chinese proverb.');
			});
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}
})();