"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
// Show the form and handle user input
//const iframe = document.getElementById("iframe") as HTMLIFrameElement;
//const clientId = "f670878d-ed8e-4020-bb82-21ba582d0d9c"; the old one
var clientId = "157dd297-447d-4592-b2d3-76b643b97132"; //the new one
var tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
var redirectUri = "https://script-lab.public.cdn.office.net";
//const redirectUri = "msal157dd297-447d-4592-b2d3-76b643b97132://auth";
var token;
(function loadMsalScript(callback) {
    var _this = this;
    var script = document.createElement("script");
    script.src = "https://alcdn.msauth.net/browser/2.17.0/js/msal-browser.min.js";
    script.onload = function () { return __awaiter(_this, void 0, void 0, function () { return __generator(this, function (_a) {
        switch (_a.label) {
            case 0: return [4 /*yield*/, getTokenWithMSAL()];
            case 1: return [2 /*return*/, (token = _a.sent())];
        }
    }); }); };
    script.onerror = function () { return console.error("Failed to load MSAL.js"); };
    document.head.appendChild(script);
})();
function showForm() {
    return __awaiter(this, void 0, void 0, function () {
        var formHtml, form;
        return __generator(this, function (_a) {
            formHtml = "<div id=\"filterForm\" >\n    <label for=\"client\" > Client: </label>\n    <input type=\"text\" id=\"client\" name=\"client\"><br><br>\n    <label for=\"affaire\" > Affaire: </label>\n    <input type =\"text\" id =\"affaire\" name=\"affaire\"><br><br>\n    <label for=\"nature\" > Nature: </label>\n    <input type =\"text\" id =\"nature\" name=\"nature\"><br><br>\n    <label for=\"date\" > Date: </label>\n    <input type=\"date\" id=\"date\" name=\"date\"><br><br>\n    <button onclick=\"filter()\"> Filter Table</button>\n  </div>";
            form = document.getElementById("form");
            if (form)
                form.innerHTML = formHtml;
            return [2 /*return*/];
        });
    });
}
function filter() {
    return __awaiter(this, void 0, void 0, function () {
        function getArray(value) {
            var array = value.split(",");
            return array.filter(function (el) { return el; });
        }
        var client, matter, nature, date, criteria;
        return __generator(this, function (_a) {
            client = document.getElementById("client");
            matter = document.getElementById("affaire");
            nature = document.getElementById("nature");
            date = document.getElementById("date");
            criteria = {
                client: getArray(client.value) || undefined,
                matter: getArray(matter.value) || undefined,
                nature: getArray(nature.value) || undefined,
                date: getArray(date.value) || undefined
            };
            filterTable(criteria);
            return [2 /*return*/];
        });
    });
}
// Filter the Excel table based on form data
function filterTable(criteria) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Excel.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                        var sheet, table, range;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    sheet = context.workbook.worksheets.getActiveWorksheet();
                                    table = sheet.tables.getItem("LivreJournal");
                                    table.autoFilter.clearCriteria();
                                    if (criteria.client.length > 0)
                                        table.columns.getItemAt(0).filter.applyValuesFilter(criteria.client);
                                    if (criteria.matter.length > 0)
                                        table.columns.getItemAt(1).filter.applyValuesFilter(criteria.matter);
                                    if (criteria.nature.length > 0)
                                        table.columns.getItemAt(2).filter.applyValuesFilter(criteria.nature);
                                    if (criteria.date.length > 0)
                                        table.columns.getItemAt(3).filter.applyValuesFilter(criteria.date);
                                    range = table.getDataBodyRange().getVisibleView();
                                    range.load("values");
                                    return [4 /*yield*/, context.sync()];
                                case 1:
                                    _a.sent();
                                    //await createWordDocument(range.values);
                                    uploadWordDocument(range.values, criteria.client.join(",") + "_Invoice" + "THEDATE");
                                    return [2 /*return*/];
                            }
                        });
                    }); })];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/*
// Create a new Word document based on a template and populate it with filtered data
async function createWordDocument(filtered: any[][]) {
  return console.log("filtered = ", filtered);

  await Word.run(async (context) => {
    const templateUrl = "https://your-onedrive-path/template.docx";
    const newDoc = context.application.createDocument(templateUrl);
    await context.sync();

    const table = newDoc.body.tables.getFirst();
    //const filteredData = await getFilteredData();

    //filtered.forEach(el) => {
    // table.(index + 1, row);
    //});

    await context.sync();

    const saveUrl = "https://your-onedrive-path/newDocument.docx";
    await newDoc.saveAs(saveUrl);
  });
}*/
// Get filtered data from the Excel table
function getTokenWithMSAL() {
    // MSAL configuration
    var msalConfig = {
        auth: {
            clientId: clientId, // Replace with your Azure AD app's client ID
            redirectUri: redirectUri
        }
    };
    //@ts-ignore
    var msalInstance = new msal.PublicClientApplication(msalConfig);
    return checkExistingAuthContext();
    // Function to check existing authentication context
    function checkExistingAuthContext() {
        return __awaiter(this, void 0, void 0, function () {
            var account;
            return __generator(this, function (_a) {
                try {
                    account = msalInstance.getAllAccounts()[0];
                    if (account) {
                        return [2 /*return*/, acquireTokenSilently(account)];
                    }
                    else {
                        return [2 /*return*/, loginAndGetToken()];
                        //openLoginWindow()
                        //return officeToken()
                        //return getTokenWithSSO('minabibawi@gmail.com')
                        //return credentitalsToken()
                    }
                }
                catch (error) {
                    console.error("Error checking auth context:", error);
                    return [2 /*return*/, undefined];
                }
                return [2 /*return*/];
            });
        });
    }
    function credentitalsToken() {
        return __awaiter(this, void 0, void 0, function () {
            var msalConfig, cca, tokenRequest, response, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        msalConfig = {
                            auth: {
                                clientId: clientId,
                                authority: "https://login.microsoftonline.com/".concat(tenantId),
                                //clientSecret: clientSecret,
                            }
                        };
                        cca = new msal.application.ConfidentialClientApplication(msalConfig);
                        tokenRequest = {
                            scopes: ["Files.ReadWrite"],
                        };
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, cca.acquireTokenByClientCredential(tokenRequest)];
                    case 2:
                        response = _a.sent();
                        return [2 /*return*/, response.accessToken];
                    case 3:
                        error_1 = _a.sent();
                        console.log('Error acquiring Token: ', error_1);
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    }
    function officeToken() {
        return __awaiter(this, void 0, void 0, function () {
            var token;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, OfficeRuntime.auth.getAccessToken()];
                    case 1:
                        token = _a.sent();
                        debugger;
                        return [2 /*return*/];
                }
            });
        });
    }
    function getTokenWithSSO(email) {
        return __awaiter(this, void 0, void 0, function () {
            var msalConfig, response, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        msalConfig = {
                            auth: {
                                clientId: clientId,
                                authority: "https://login.microsoftonline.com/".concat(tenantId),
                                redirectUri: redirectUri,
                                navigateToLoginRequestUrl: true,
                            },
                            cache: {
                                cacheLocation: "sessionStorage",
                                storeAuthStateInCookie: true
                            }
                        };
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, msalInstance.ssoSilent({
                                scopes: ["Files.ReadWrite.All"],
                                //scopes: ["https://graph.microsoft.com/.default"],
                                loginHint: email // Forces MSAL to recognize the signed-in user
                            })];
                    case 2:
                        response = _a.sent();
                        console.log("Token acquired via SSO:", response.accessToken);
                        return [2 /*return*/, response.accessToken];
                    case 3:
                        error_2 = _a.sent();
                        console.error("SSO silent authentication failed:", error_2);
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    }
    function openLoginWindow() {
        var loginUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=".concat(clientId, "&response_type=token&redirect_uri=").concat(redirectUri, "&scope=https://graph.microsoft.com/.default");
        // Open in a new window (only works if triggered by user action)
        var authWindow = window.open(loginUrl, "_blank", "width=500,height=600");
        if (!authWindow) {
            console.error("Popup blocked! Please allow popups.");
        }
    }
    // Function to handle login and acquire token
    function loginAndGetToken() {
        return __awaiter(this, void 0, void 0, function () {
            // Function to handle redirect response
            function handleRedirectResponse() {
                return __awaiter(this, void 0, void 0, function () {
                    var authResult, error_4;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                _a.trys.push([0, 2, , 3]);
                                return [4 /*yield*/, msalInstance.handleRedirectPromise()];
                            case 1:
                                authResult = _a.sent();
                                if (authResult && authResult.accessToken) {
                                    console.log("Access token:", authResult.accessToken);
                                    return [2 /*return*/, authResult.accessToken];
                                }
                                return [3 /*break*/, 3];
                            case 2:
                                error_4 = _a.sent();
                                console.error("Redirect handling error:", error_4);
                                return [3 /*break*/, 3];
                            case 3: return [2 /*return*/, undefined];
                        }
                    });
                });
            }
            var loginRequest, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        loginRequest = {
                            scopes: ["Files.ReadWrite"] // OneDrive scopes
                        };
                        return [4 /*yield*/, msalInstance.loginRedirect(loginRequest)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, handleRedirectResponse()];
                    case 2:
                        error_3 = _a.sent();
                        console.error("Login error:", error_3);
                        return [2 /*return*/, undefined];
                    case 3: return [2 /*return*/];
                }
            });
        });
    }
    // Function to get access token silently
    function acquireTokenSilently(account) {
        return __awaiter(this, void 0, void 0, function () {
            var tokenRequest, tokenResponse, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        tokenRequest = {
                            scopes: ["Files.ReadWritel"], // OneDrive scopes
                            account: account
                        };
                        return [4 /*yield*/, msalInstance.acquireTokenSilent(tokenRequest)];
                    case 1:
                        tokenResponse = _a.sent();
                        if (tokenResponse && tokenResponse.accessToken) {
                            console.log("Access token (silent):", tokenResponse.accessToken);
                            return [2 /*return*/, tokenResponse.accessToken];
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        error_5 = _a.sent();
                        console.error("Token acquisition error:", error_5);
                        return [2 /*return*/, loginAndGetToken()];
                    case 3: return [2 /*return*/, undefined];
                }
            });
        });
    }
}
/*
async function authenticateUser() {
  const clientSecret = "Inl8Q~jhDg8qQ5jrhBTuQBCQbGdkHmcQLpMqEcTQ";
  const secretID = "ad646418-c15f-44b1-90cc-0af31238d1e6";
  const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  //return noLogin();
  return withLogin();

  async function noLogin() {
    const params = new URLSearchParams();
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("scope", "https://graph.microsoft.com/.default");
    params.append("grant_type", "client_credentials");
    console.log("params =", params);
    try {
      const response = await fetch(tokenEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params
      });

      const data = await response.json();
      if (data.access_token) {
        console.log("✅ Access Token Retrieved:", data.access_token);
        return data.access_token;
      } else {
        throw new Error("❌ Failed to retrieve access token: " + JSON.stringify(data));
      }
    } catch (error) {
      console.error("Error getting access token:", error);
    }
  }

  function withLogin() {
    const scopes = "Files.ReadWrite";
    const redirectUri = encodeURI("https://login.microsoftonline.com/common/oauth2/nativeclient");
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=${scopes}&response_mode=fragment`;
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(authUrl, { height: 60, width: 40 }, (result) => {
        console.log("result = ", result);
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            const token = new URLSearchParams(args.message.split("#")[1]).get("access_token");
            dialog.close();
            debugger;
            if (token) resolve(token);
            else reject("Authentication failed");
          });
        } else {
          reject("Failed to open login dialog.");
        }
      });
    });
  }
}*/
/*
async function getAccessToken() {

  const scopes = ["Files.ReadWrite", "User.Read"];
  const redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient";

  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=${scopes.join(
    "%20"
  )}&response_mode=fragment`;

  // Open authentication window
  const authWindow = window.open(authUrl, "_blank");

  return new Promise((resolve, reject) => {
    const checkAuth = setInterval(() => {
      try {
        if (authWindow.location.hash) {
          clearInterval(checkAuth);
          const token = new URLSearchParams(authWindow.location.hash.substring(1)).get("access_token");
          authWindow.close();
          resolve(token);
        }
      } catch (e) {}
    }, 1000);
  });
}*/
function uploadWordDocument(filtered, fileName) {
    return __awaiter(this, void 0, void 0, function () {
        var accessToken, wordContent, oneDriveUrl, response, _a, _b, _c;
        return __generator(this, function (_d) {
            switch (_d.label) {
                case 0: return [4 /*yield*/, getTokenWithMSAL()];
                case 1:
                    accessToken = _d.sent();
                    if (accessToken) {
                        console.log("Successfully retrieved token:", accessToken);
                        Office.context.ui.messageParent("Access token: ".concat(accessToken));
                    }
                    else {
                        console.log("Failed to retrieve token.");
                    }
                    if (!accessToken)
                        return [2 /*return*/, console.log("No access token")];
                    wordContent = "UEsDBBQAAAAIA...";
                    oneDriveUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/Documents/".concat(fileName, ".docx:/content");
                    return [4 /*yield*/, fetch(oneDriveUrl, {
                            method: "PUT",
                            headers: {
                                Authorization: "Bearer ".concat(accessToken),
                                "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            },
                            body: atob(wordContent) // Convert base64 to binary
                        })];
                case 2:
                    response = _d.sent();
                    if (!response.ok) return [3 /*break*/, 3];
                    console.log("File uploaded successfully!");
                    return [3 /*break*/, 5];
                case 3:
                    _b = (_a = console).error;
                    _c = ["Upload failed"];
                    return [4 /*yield*/, response.text()];
                case 4:
                    _b.apply(_a, _c.concat([_d.sent()]));
                    _d.label = 5;
                case 5: return [2 /*return*/];
            }
        });
    });
}
// Show the form when the script is run
Office.onReady(function () { return showForm(); });
//uploadWordDocument();
