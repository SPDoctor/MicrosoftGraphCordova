// For an introduction to the Blank template, see the following documentation:
// http://go.microsoft.com/fwlink/?LinkID=397704
// To debug code on page load in Ripple or on Android devices/emulators: launch your app, set breakpoints, 
// and then run "window.location.reload()" in the JavaScript Console.
(function () {
	"use strict";
	var tenantName = 'flosim';
	var authority = "https://login.windows.net/" + tenantName + ".onmicrosoft.com";
	var resourceUrl = 'https://graph.microsoft.com/';
	var appId = "92f98787-c980-4c15-9be0-348ba4244408";
	var redirectUrl = "http://localhost:4400/www/index.html";
	var authContext;
	var output;

	document.addEventListener('deviceready', onDeviceReady.bind(this), false);

	function onDeviceReady() {
		// Handle the Cordova pause and resume events
		document.addEventListener('pause', onPause.bind(this), false);
		document.addEventListener('resume', onResume.bind(this), false);
		document.getElementById("logout").addEventListener('click', onLogout.bind(this), false);
		document.getElementById("loadToken").addEventListener('click', onLoadToken.bind(this), false);
		output = document.getElementById("output");

		// TODO: Cordova has been loaded. Perform any initialization that requires Cordova here.
	};

	function fail(err) {
		output.innerHTML = err.message;
	}

	function onLoadToken() {
		getAccessToken(resourceUrl, appId, redirectUrl, function (response) {
			var message = "";
			message += "Access token: " + response.accessToken;
			message += "<br />\r\n";
			message += "Decoded token: " + decodeJWT(response.accessToken);
			message += "<br />\r\n";
			message += "Token will expire on: " + response.expiresOn;
			message += "<br />\r\n";
			output.innerHTML = message;

			var endPointUri = "https://graph.microsoft.com/v1.0/";
			var requestUri = endPointUri + 'me/contacts?$top=20';
			var bearerToken = "Bearer " + response.accessToken;
			var xhr = new XMLHttpRequest();
			xhr.open('GET', requestUri);
			xhr.setRequestHeader("Authorization", bearerToken);
			xhr.setRequestHeader("Accept", "application/json;odata.metadata=minimal");
			xhr.onload = function () {
				if (xhr.status === 200) {
					var response = JSON.parse(xhr.responseText);
					var contacts = "";
					for (var i = 0; i < response.value.length; i++) {
						contacts += "<li>" + response.value[i].displayName + "</li>";
					};
					document.getElementById("contacts").innerHTML = contacts;
				}
				else {
					output.innerHTML += '<br />Request failed.  Returned status of ' + xhr.status;
				}
			};
			xhr.send();
		}, fail);
	};

	function onLogout() {
		if (authContext) authContext.tokenCache.clear();
	}

	function getAccessToken1(resourceUrl, appId, redirectUrl, success, fail) {
		try {
			// can try to use acquireTokenSilent here if already have a token
			authContext.tokenCache.readItems().then(function (cacheItems) {
				if (cacheItems && cacheItems.length > 0) {
					success(cacheItems[0]);
					return;
				}
				authContext.acquireTokenAsync(resourceUrl, appId, redirectUrl).then(function (authResponse) {
					success(authResponse);
				}, fail);
			});
		} catch (ex) {
			fail(ex.message);
		}
	}

	function getAccessToken(resourceUrl, appId, redirectUrl, success, fail) {
		try {
			if (!authContext) {
				Microsoft.ADAL.AuthenticationContext.createAsync(authority).then(function (context) {
					authContext = context;
					getAccessToken1(resourceUrl, appId, redirectUrl, success, fail);
				}, fail);
			}
			else getAccessToken1(resourceUrl, appId, redirectUrl, success, fail);
		} catch (ex) {
			fail(ex.message);
		}
	}

	function decodeJWT(encodedJWT) {
		var decodedJWT = "";
		try {
			var sections = encodedJWT.split(".");
			decodedJWT += "<br /><b>Header:</b> " + atob(sections[0]);
			decodedJWT += "<br /><b>Payload:</b> " + atob(sections[1]);
			decodedJWT += "<br /><b>Signature:</b> " + sections[2];
			decodedJWT += "<br />";
		} catch (ex) {
			decodedJWT += "<br />Error: " + ex.message;
		}
		return decodedJWT;
	}

	function onPause() {
		// TODO: This application has been suspended. Save application state here.
	};

	function onResume() {
		// TODO: This application has been reactivated. Restore application state here.
	};
})();