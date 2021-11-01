const FileList = document.getElementById('FileList');
let spFiles = [];

const loadFiles = async () => {
    try {
        //const res = await fetch('https://hp-api.herokuapp.com/api/characters');
        const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
        spFiles = await graphData.json();
        displayFiles(spFiles);
    } catch (err) {
        console.error(err);
    }
};

const displayFiles = (files) => {
    const htmlString = files
        .map((file) => {
            return `
            <li class="file">
                <h2>${file.name}</h2>
                <h2>${file.id}</h2>
            </li>
        `;
        })
        .join('');
    charactersList.innerHTML = htmlString;
};
app.get('/getuserdata/:searchTerm', async function (req, res, next) {
	const graphToken = req.get('access_token');

	console.log(req.params.searchTerm);

	const itemNames = [];
	res.send(itemNames);

	// Minimize the data that must come from MS Graph by specifying only the property we need ("name")
	// and only the top 10 folder or file names.
	// Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
	// a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
	// sanitized so that it cannot be used in a Response header injection attack.

	const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");

	//const SPpathwSearch = "/groups/{af65c412-5d67-4611-abe8-ec0a68e8b7f3}/drive/root/search(q='Category')";

//	const SPpathwSearch = `/groups/{af65c412-5d67-4611-abe8-ec0a68e8b7f3}/drive/root/search(q=${req.params.searchTerm})`;
//	const graphData = await getGraphData(graphToken, SPpathwSearch, '?select=name,id,webUrl');

	// If Microsoft Graph returns an error, such as invalid or expired token,
	// there will be a code property in the returned object set to a HTTP status (e.g. 401).
	// Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
	if (graphData.code) {
		next(createError(graphData.code, 'Microsoft Graph error ' + JSON.stringify(graphData)));
	} else {
		// MS Graph data includes OData metadata and eTags that we don't need.
		// Send only what is actually needed to the client: the item names.
		const itemNames = [];
		const oneDriveItems = graphData['value'];
		for (let item of oneDriveItems) {
			itemNames.push(item['name']); //-,item['id'],item['webUrl'],item['thumbnails']
		}
		res.send(itemNames);
	}
});
function searchForData() {
	return __awaiter(this, void 0, void 0, function () {
		var jSearch;
		return __generator(this, function (_a) {
			switch (_a.label) {
				case 0:
					console.log('searchForData');
					jSearch = document.getElementById('tSearch').value;
					return [4 /*yield*/, getGraphData(jSearch)];
				case 1:
					_a.sent();
					return [2 /*return*/];
			}
		});
	});
}
var retryGetAccessToken = 0;
var searchTerm = '';
function getGraphData() {
	return __awaiter(this, void 0, void 0, function () {
		var bootstrapToken, exchangeResponse, mfaBootstrapToken, exception_1;
		return __generator(this, function (_a) {
			switch (_a.label) {
				case 0:
					console.log('getGraphData');
					searchTerm = document.getElementById('tSearch').value;
					console.log(searchTerm);
					_a.label = 1;
				case 1:
					_a.trys.push([1, 7, , 8]);
					return [4 /*yield*/, OfficeRuntime.auth.getAccessToken({
						allowSignInPrompt: true,
						allowConsentPrompt: true,
						forMSGraphAccess: true,
					})];
				case 2:
					bootstrapToken = _a.sent();
					return [4 /*yield*/, getGraphToken(bootstrapToken)];
				case 3:
					exchangeResponse = _a.sent();
					if (!exchangeResponse.claims) return [3 /*break*/, 6];
					return [4 /*yield*/, OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims })];
				case 4:
					mfaBootstrapToken = _a.sent();
					return [4 /*yield*/, getGraphToken(mfaBootstrapToken)];
				case 5:
					exchangeResponse = _a.sent();
					_a.label = 6;
				case 6:
					if (exchangeResponse.error) {
						// AAD errors are returned to the client with HTTP code 200, so they do not trigger
						// the catch block below.
						handleAADErrors(exchangeResponse);
					}
					else {
						// For debugging:
						// showMessage("ACCESS TOKEN: " + JSON.stringify(exchangeResponse.access_token));
						// makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
						// in the .fail callback of that call, not in the catch block below.
						console.log('calling makeGraphAPICall', jSearch);
						makeGraphApiCall(exchangeResponse.access_token); //, searchTerm);
					}
					return [3 /*break*/, 8];
				case 7:
					exception_1 = _a.sent();
					// The only exceptions caught here are exceptions in your code in the try block
					// and errors returned from the call of `getAccessToken` above.
					if (exception_1.code) {
						handleClientSideErrors(exception_1);
					}
					else {
						showMessage('EXCEPTION: ' + JSON.stringify(exception_1));
					}
					return [3 /*break*/, 8];
				case 8: return [2 /*return*/];
			}
		});
	});
}
function getGraphToken(bootstrapToken) {
	return __awaiter(this, void 0, void 0, function () {
		var response;
		return __generator(this, function (_a) {
			switch (_a.label) {
				case 0: return [4 /*yield*/, $.ajax({
					type: 'GET',
					url: '/auth',
					headers: { Authorization: 'Bearer ' + bootstrapToken },
					cache: false,
				})];
				case 1:
					response = _a.sent();
					return [2 /*return*/, response];
			}
		});
	});
}
function handleClientSideErrors(error) {
	switch (error.code) {
		case 13001:
			// No one is signed into Office. If the add-in cannot be effectively used when no one
			// is logged into Office, then the first call of getAccessToken should pass the
			// `allowSignInPrompt: true` option. Since this sample does that, you should not see
			// this error.
			showMessage('No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.');
			break;
		case 13002:
			// The user aborted the consent prompt. If the add-in cannot be effectively used when consent
			// has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
			showMessage('You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.');
			break;
		case 13006:
			// Only seen in Office on the web.
			showMessage('Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.');
			break;
		case 13008:
			// Only seen in Office on the web.
			showMessage('Office is still working on the last operation. When it completes, try this operation again.');
			break;
		case 13010:
			// Only seen in Office on the web.
			showMessage("Follow the instructions to change your browser's zone configuration.");
			break;
		default:
			// For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
			// to non-SSO sign-in.
			dialogFallback();
			break;
	}
}
async function useInsertSlideApi() {
	const myFile = <HTMLInputElement> document.getElementById("file");
		const reader = new FileReader();

  reader.onload = async (event) => {
    // strip off the metadata before the base64-encoded string
    const startIndex = reader.result.toString().indexOf("base64,");
		const copyBase64 = reader.result.toString().substr(startIndex + 7);

		await PowerPoint.run(async function(ctx) {
			ctx.presentation.insertSlidesFromBase64(copyBase64, { formatting: "UseDestinationTheme" });
		// "targetSlideId"
		// "sourceSlideIds"
		ctx.sync();
    });
  };

		// read in the file as a data URL so we can parse the base64-encoded string
		reader.readAsDataURL(myfile.files[0]);
}
		function handleAADErrors(exchangeResponse) {
		// On rare occasions the bootstrap token is unexpired when Office validates it,
		// but expires by the time it is sent to AAD for exchange. AAD will respond
		// with "The provided value for the 'assertion' is not valid. The assertion has expired."
		// Retry the call of getAccessToken (no more than once). This time Office will return a
		// new unexpired bootstrap token.
		if (exchangeResponse.error_description.indexOf('AADSTS500133') !== -1 && retryGetAccessToken <= 0) {
			retryGetAccessToken++;
		getGraphData();
		}
		else {
			// For all other AAD errors, fallback to non-SSO sign-in.
			// For debugging:
			// showMessage("AAD ERROR: " + JSON.stringify(exchangeResponse));
			dialogFallback();
		}
	}
loadFiles();