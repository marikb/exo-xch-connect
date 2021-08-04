var item;
var clickEvent;


Office.initialize = function () {
	item = Office.context.mailbox.item;
	// Checks for the DOM to load using the jQuery ready function.
	$(document).ready(function () {
		// After the DOM is loaded, app-specific code can run.
		// Get the subject of the item being composed.
		getSubject();
		


	});
}

// Get the subject of the item that the user is composing.
function getSubject() {
	item.subject.getAsync(
		function (asyncResult) {
			if (asyncResult.status == Office.AsyncResultStatus.Failed) {
				write(asyncResult.error.message);
			}
			else {
				// Successfully got the subject, display it.
				var subject = asyncResult.value;
				setSubject(subject + '@MAAP@')
				//sendEmail(null)
			}
		});
}

function setSubject(newSubject) {
	var subject;

	// Customize the subject with today's date.
	subject = newSubject;

	item.subject.setAsync(
		subject,
		{ asyncContext: { var1: 1, var2: 2 } },
		function (asyncResult) {
			if (asyncResult.status == Office.AsyncResultStatus.Failed) {
				write(asyncResult.error.message);
			}
			else {
				// Successfully set the subject.
				// Do whatever appropriate for your scenario
				// using the arguments var1 and var2 as applicable.
			}
		});
}


// Write to a div with id='message' on the page.
function write(message) {
	console.log(message);
}
//The function that is called when we click on «Analyze And Send»
function sendEmail(event) {
	clickEvent = event;
	console.log(Office.context.mailbox.item)
	//If we create a new item we need to save it to draft to get item Id
	if (Office.context.mailbox.item.itemId === null || Office.context.mailbox.item.itemId == undefined) {
		Office.context.mailbox.item.saveAsync(saveItemCallBack);
		console.log('a')
	}
	else {
		console.log('b')
		var soapToGetItemData = getItemDataRequest(Office.context.mailbox.item.itemId);
		Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, itemDataCallback);
	}
}

function saveItemCallBack(result) {
	var soapToGetItemData = getItemDataRequest(result.value);
	//Make Ews request to get item info
	Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, itemDataCallback);
}

function itemDataCallback(asyncResult) {
	if (asyncResult.error != null) {
		updateAndComplete("EWS Status: " + asyncResult.error.message);
		return;
	}
	//Parse response from EWS
	console.log(asyncResult.value)
	var xmlDoc = getXMLDocParser(asyncResult.value);
	console.log(xmlDoc)


	var result = $('ResponseCode', xmlDoc)[0].textContent;
	if (result != "NoError") {
		updateAndComplete("EWS Status", "The following error code was recieved: " + result);
		return;
	}

	//Get information about attachments from response
	var attachmentsInfo = buildAttachmentsInfo(xmlDoc);
	Office.context.mailbox.item.loadCustomPropertiesAsync(function (asyncResult) {
		//Set custom properties
		var customProps = asyncResult.value;
		customProps.set("myProp", "value");
		customProps.saveAsync(function (asyncResult) {
			if (asyncResult.status == Office.AsyncResultStatus.Failed) {
				updateAndComplete(asyncResult.error.message);
				return;
			}

			modifyEmailAndSend(attachmentsInfo);
		});
	});
}

function modifyEmailAndSend(attachmentsInfo) {
	//Modify item body. Add to the end of item information about attachments
	Office.context.mailbox.item.body.getAsync("html", { asyncContext: "This is passed to the callback" }, function (result) {
		var newText = result.value + "<br>" + attachmentsInfo;
		Office.context.mailbox.item.body.setAsync(newText, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
			if (asyncResult.status != Office.AsyncResultStatus.Succeeded) {
				statusUpdate("Couldn't modify body");
				return;
			}
			//When we changed and saved message body we need to get a new Change key to send the message
			Office.context.mailbox.item.saveAsync(function (result) {
				var soapToGetItemData = getItemDataRequest(result.value);
				Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, function (asyncResult) {
					if (asyncResult.error != null) {
						updateAndComplete("EWS Status: " + asyncResult.error.message);
						return;
					}

					var xmlDoc = getXMLDocParser(asyncResult.value);
					var changeKey = $('ItemId', xmlDoc)[0].getAttribute("ChangeKey");
					//Send the message
					var soapToSendItem = getSendItemRequest(result.value, changeKey);
					Office.context.mailbox.makeEwsRequestAsync(soapToSendItem, function (asyncResult) {
						if (asyncResult.error != null) {
							statusUpdate("EWS Status: " + asyncResult.error.message);
							return;
						}

						Office.context.mailbox.item.close();
						clickEvent.completed();
					});
				});
			});

		});
	});
}
//Office.initialize = function () {
//	item = Office.context.mailbox.item
//	$(document).ready(function () {
//		// After the DOM is loaded, app-specific code can run.
//		// Get the subject of the item being composed.
//		console.log(getSubject());
//	});
//}

//// Helper function to add a status message to the info bar.
//function statusUpdate(icon, text) {
//  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
//    type: "informationalMessage",
//    icon: icon,
//    message: text,
//    persistent: false
//  });
//}

//function defaultStatus(event) {
//  statusUpdate("icon16" , "Hello World!");
//}

//function insertDefaultGist(event) {
//	Office.context.mailbox.item.subject.get
//	console.log(getSubject() + '@MyPlaceHolder@')
//	//var subj = await Office.context.mailbox.item.subject.getAsync()
//	//console.log(subj)
//	//await Office.context.mailbox.item.subject.setAsync(subj + "@placeholder@")
//	//console.log(await Office.context.mailbox.item.subject.getAsync())
//	//console.log(event)
//	//console.log(Office)
//}

//function getSubject() {
//	item.subject.getAsync(
//		function (asyncResult) {
//			if (asyncResult.status == Office.AsyncResultStatus.Failed) {
//				return null;
//			}
//			else {
//				// Successfully got the subject, display it.
//				return asyncResult.value;
//			}
//		});
//}

//var addPlaceholder = async (s) => {
//	var placeholder = '@MyPlaceHolder@'
//	return getSubject() + placeholder
//}



///---------------------------------------------------------------------------------------------------

//Ews request to get item info
function getItemDataRequest(itemId) {
	var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
		'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
		'               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
		'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
		'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
		'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
		'  <soap:Header>' +
		'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
		'  </soap:Header>' +
		'  <soap:Body>' +
		'    <GetItem' +
		'                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
		'                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
		'      <ItemShape>' +
		'        <t:BaseShape>IdOnly</t:BaseShape>' +
		'        <t:AdditionalProperties>' +
		'            <t:FieldURI FieldURI="item:Attachments" /> ' +
		'        </t:AdditionalProperties> ' +
		'      </ItemShape>' +
		'      <ItemIds>' +
		'        <t:ItemId Id="' + itemId + '"/>' +
		'      </ItemIds>' +
		'    </GetItem>' +
		'  </soap:Body>' +
		'</soap:Envelope>';

	return soapToGetItemData;
}

//Ews request to send the modified item
function getSendItemRequest(itemId, changeKey) {
	var soapSendItemRequest = '<?xml version="1.0" encoding="utf-8"?>' +
		'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
		'               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
		'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
		'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
		'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
		'  <soap:Header>' +
		'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
		'  </soap:Header>' +
		'  <soap:Body> ' +
		'    <SendItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
		'              SaveItemToFolder="true"> ' +
		'      <ItemIds> ' +
		'        <t:ItemId Id="' + itemId + '" ChangeKey="' + changeKey + '" /> ' +
		'      </ItemIds> ' +
		'      <m:SavedItemFolderId>' +
		'         <t:DistinguishedFolderId Id="sentitems" />' +
		'      </m:SavedItemFolderId>' +
		'    </SendItem> ' +
		'  </soap:Body> ' +
		'</soap:Envelope> ';
	return soapSendItemRequest;
}


function buildAttachmentsInfo(xmlDoc) {
	var attachmentsInfo = "You have no any attachments.";
	if ($('HasAttachments', xmlDoc).length == 0) {
		return attachmentsInfo;
	}

	var attachSeparator = "--------------------------------------------- <br>";
	if ($('HasAttachments', xmlDoc)[0].textContent == "true") {
		attachmentsInfo = "";
		var childNodes = $('Attachments', xmlDoc)[0].childNodes;
		childNodes.forEach(function (fileAttachmentItem, fileAttachmentIndex) {
			fileAttachmentItem.childNodes.forEach(function (item, index) {
				if (item.tagName.includes("AttachmentId")) {
					attachmentsInfo = attachmentsInfo.concat(item.tagName.replace("t:", "") + ': ' + item.getAttribute("Id") + "<br>");
					return;
				}

				attachmentsInfo = attachmentsInfo.concat(item.tagName.replace("t:", "") + ': ' + item.textContent + "<br>");
			});

			attachmentsInfo = attachmentsInfo.concat(attachSeparator);
		});
	}

	return attachmentsInfo;
}

function updateAndComplete(text) {
	//Notify in UI about status
	Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
		type: "informationalMessage",
		message: text,
		icon: "default_16",
		persistent: false
	});

	clickEvent.completed();
}

function getXMLDocParser(response) {
	var xmlDoc;
	if (window.DOMParser) {
		var parser = new DOMParser();
		xmlDoc = parser.parseFromString(response, "text/xml");
	}
	else // Older Versions of Internet Explorer
	{
		xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
		xmlDoc.async = false;
		xmlDoc.loadXML(response);
	}
	return xmlDoc;
}
