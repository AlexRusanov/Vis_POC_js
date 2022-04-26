// If the add-in is running in Internet Explorer, the code must add support
// for Promises.
if (!window.Promise) {
  window.Promise = Office.Promise;
}

const requestObj = {
  scopes: ["https://graph.microsoft.com/User.Read"],
};
Office.initialize = function () {
  if (Office.context.ui.messageParent) {
    userAgentApp.handleRedirectCallback(authCallback);

    if (localStorage.getItem("loggedIn") === "yes") {
      userAgentApp.acquireTokenRedirect(requestObj);
    } else {
      userAgentApp.loginRedirect(requestObj);
    }
  }
};

const msalConfig = {
  auth: {
    clientId: "462f58e1-390a-4d89-8594-64b969f63bf0", //This is your client ID
    authority: "https://login.microsoftonline.com/common",
    redirectURI: "https://localhost:3000/dialog.html",
    navigateToLoginRequestUrl: false,
    response_type: "code", // access_token
  },
  cache: {
    cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true, // Recommended to avoid certain IE/Edge issues.
  },
};

const userAgentApp = new Msal.UserAgentApplication(msalConfig);

function authCallback(error, response) {
  if (error) {
    console.log(error);
    Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
  } else {
    if (response.tokenType === "id_token") {
      console.log(response.idToken.rawIdToken);
      localStorage.setItem("loggedIn", "yes");
    } else {
      console.log("token type is:" + response.tokenType);
      Office.context.ui.messageParent(JSON.stringify({ status: "success", result: response.accessToken }));
    }
  }
}
