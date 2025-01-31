Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Excel-specific initialization code goes here
    console.log("Excel is ready!");

    loadMsalScript();
  }
});


function loadMsalScript() {
  var token;
  const script = document.createElement("script");
  script.src = "https://alcdn.msauth.net/browser/2.17.0/js/msal-browser.min.js";
  script.onload = async () => (token = await getTokenWithMSAL());
  script.onerror = () => console.error("Failed to load MSAL.js");
  document.head.appendChild(script);
  token ? console.log('Got token', token) : console.log('could not retrieve the token');
};


function selectForm(id: string) {
  showForm(id)
}

async function showForm(id?: string) {

  const form = document.getElementById("form") as HTMLDivElement;
  form.innerHTML = '';
  if (!form) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem('LivreJournal');
    const header = table.getHeaderRowRange()
    header.load('values');
    table.columns.load('items');
    const dataBodyRange = table.getDataBodyRange()
    dataBodyRange.load('values');
    await context.sync();
    const title = header.values[0];

    if (id === 'entry') await addingEntry(table.columns.items, title, dataBodyRange.values);
    else if (id === 'invoice') await invoice(title);
  });

  async function invoice(title: any[]) {
    const inputs = await Promise.all(insertInputsAndLables([0, 1, 2, 3]));
    form.innerHTML += `<button onclick="generateInvoice()"> Filter Table</button>`;
    inputs.forEach(input => input?.addEventListener('focusout', async () => await inputOnChange(input), { passive: true }));

    function insertInputsAndLables(indexes: number[]): Promise<(HTMLInputElement | undefined)>[] {
      const id = 'input';
      return indexes.map(async index => {
        if (!index) return;
        const input = document.createElement('input');
        input.type = "text"
        if (index === 3) input.type = 'date';
        input.id = id + index.toString();
        input.name = input.id;
        input.dataset.index = index.toString();
        input.setAttribute('list', input.id + 's'),
          input.autocomplete = "on"

        const label = document.createElement('label');
        label.htmlFor = input.id;
        label.innerText = title[index];

        form.appendChild(label);
        form.appendChild(input);
        if (index === 0) createDataList(input?.id, await getUniqueValues(index));//We create a unique values dataList for the 'Client' input
        return input
      });
    };

    async function inputOnChange(input: HTMLInputElement, unfilter: boolean = false) {
      const index = Number(input.dataset.index);

      if (index === 0) unfilter = true;//If this is the 'Client' column, we remove any filter from the table;

      //We filter the table accordin to the input's value and return the visible cells
      const visibleCells = await filterTable(undefined, [{ column: index, value: getArray(input.value) }], unfilter);

      if (visibleCells.length < 1) return alert('There are no visible cells in the filtered table');

      //We create (or update) the unique values dataList for the next input 
      const nextInput = getNextInput(input);
      if (!nextInput) return;
      createDataList(nextInput?.id || '', await getUniqueValues(Number(nextInput.dataset.index), visibleCells));


      function getNextInput(input: HTMLInputElement) {
        let nextInput: Element | null = input.nextElementSibling;
        while (nextInput?.tagName !== 'INPUT' && nextInput?.nextElementSibling) {
          nextInput = nextInput.nextElementSibling
        };

        return nextInput as HTMLInputElement
      }

      if (index === 1) {
        //!Need to figuer out how to create a multiple choice input for nature
        const nature = new Set((await filterTable(undefined, undefined, false)).map(row => row[index]));
        nature.forEach(el => form.appendChild(createCheckBox(undefined, el)));
      }

    };
  }

  async function addingEntry(columns: Excel.TableColumn[], title: any[], dataBodyRange: any[][]) {
    await filterTable(undefined, undefined, true);

    for (const col of columns) {
      const i = columns.indexOf(col);
      if (![4, 7].includes(i)) form.appendChild(createLable(i));//We exclued the labels for "Total Time" and for "Year"
      form.appendChild(await createInput(i));
    };

    const inputs = Array.from(document.getElementsByTagName('input'));
    inputs
      .filter(input => Number(input?.dataset.index) < 2)
      .forEach(input => input?.addEventListener('change', async () => await onFoucusOut(input), { passive: true }));

    inputs
      .filter(input => [4, 7].includes(Number(input?.dataset.index)))
      .forEach(input => input.style.display = 'none');



    async function onFoucusOut(input: HTMLInputElement) {
      debugger
      const i = Number(input.dataset.index);
      const criteria = [{ column: i, value: getArray(input.value) }];
      let unfilter = false;
      if (i === 0) unfilter = true;
      await filterTable(undefined, criteria, unfilter);
      if (i < 1)
        createDataList('input' + String(i + 1), await getUniqueValues(i + 1));
    }

    form.innerHTML += `<button onclick="addEntry()"> Ajouter </button>`;


    function createLable(i: number) {
      const label = document.createElement('label');
      label.htmlFor = 'input' + i.toString();
      label.innerHTML = title[i] + ':';
      return label
    }

    async function createInput(i: number) {
      const input = document.createElement('input');
      const id = 'input';
      input.name = id + i.toString();
      input.id = input.name;
      input.autocomplete = "on";
      input.dataset.index = i.toString();
      input.type = 'text';
      if ([8, 9, 10].includes(i))
        input.type = 'number';
      else if (i === 3)
        input.type = 'date';
      else if ([5, 6].includes(i))
        input.type = 'time';
      else if ([0, 1, 2, 11, 12, 13, 16].includes(i)) {
        //We add a dataList for those fields
        input.setAttribute('list', input.id + 's');
        createDataList(input.id, await getUniqueValues(i, dataBodyRange));
      }

      return input

    }

  }



  function createCheckBox(input: HTMLInputElement | undefined, id: string = '') {
    if (!input) input = document.createElement('input');
    input.type = 'checkbox';
    input.id += id;


    return input


  };

  async function getUniqueValues(index: number, visible?: any[][], tableName: string = 'LivreJournal') {
    if (!visible) visible = await filterTable(tableName, undefined, false);
    return Array.from(new Set(visible.map(row => row[index])))
  };



  /**
   * Creates a dataList with the provided id from the unique values of the column which index is passed as parameter
   * @param {string} id - the id of the dataList that will be created
   * @param {number} index - the index of the column from which the unique values of the datalist will be retrieved
   * 
  */

  function createDataList(id: string, uniqueValues: any[]) {
    //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
    if (!id || !uniqueValues || uniqueValues.length < 1) return;
    id += 's';

    // Create a new datalist element
    let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === id);
    if (dataList) dataList.remove();
    dataList = document.createElement('datalist');
    //@ts-ignore
    dataList.id = id;
    // Append options to the datalist
    uniqueValues.forEach(option => {
      const optionElement = document.createElement('option');
      optionElement.value = option;
      dataList?.appendChild(optionElement);
    });
    // Attach the datalist to the body or a specific element
    document.body.appendChild(dataList);
  };

}

/**
 * Filters the Excel table based on a criteria
 * @param {[[number, string[]]]} criteria - the first element is the column index, the second element is the values[] based on which the column will be filtered
 */
async function filterTable(tableName: string = 'LivreJournal', criteria?: { column: number, value: string[] }[], clearFilter: boolean = false) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem(tableName);

    if (clearFilter) table.autoFilter.clearCriteria();

    if (!criteria) return await getVisible();

    criteria.forEach(column => filterColumn(column.column, column.value));

    return await getVisible();

    function filterColumn(index: number, filter: string[]) {
      if (!index || !filter) return;
      table.columns.getItemAt(index).filter.applyValuesFilter(filter)
    }

    async function getVisible() {
      const visible = table.getDataBodyRange().getVisibleView();
      visible.load('values');
      await context.sync();
      return visible.values
    }
  });
}

/**
 * Converts the ',' separated text in the input into an array
 * @param value 
 * @returns {string[]}
 */
function getArray(value: string): string[] {
  const array =
    value.replaceAll(', ', ',')
      .replaceAll(' ,', ',')
      .split(',');
  return array.filter((el) => el);
}

async function generateInvoice() {
  const inputs = Array.from(document.getElementsByName('input')) as HTMLInputElement[];
  if (!inputs) return;
  const date = new Date();
  const fileName = "Test_Facture_" + inputs.find(input => Number(input.dataset.index) === 0)?.value || 'CLIENT_' + [date.getDay().toString(), (date.getMonth() + 1).toString(), date.getFullYear.toString()].join('-') + '_' + date.getHours().toString() + date.getMinutes().toString() + '.docx';

  const visible = await filterTable(undefined, undefined, false);

  await uploadWordDocument(visible, fileName);

}

async function uploadWordDocument(filtered: any[][], fileName: string) {
  //const accessToken = await authenticateUser();
  const accessToken = await getTokenWithMSAL();
  if (accessToken) {
    console.log("Successfully retrieved token:", accessToken);
    alert(`Access token: ${accessToken}`);
    //Office.context.ui.messageParent(`Access token: ${accessToken}`);
  } else {
    return console.log("Failed to retrieve token.");
  }

  const path = "Legal\\Mon Cabinet d'Avocat\\Comptabilité\\Factures\\"
  const templatePath = path + "FactureTEMPLATE [NE PAS MODIFIDER].dotm";
  const newPath = path + `Clients\\${fileName}`;

  createWordDocumentFromTemplate(templatePath, newPath, accessToken)

  return
  // Sample Word document content (base64 encoded DOCX)
  const wordContent = "UEsDBBQAAAAIA...";

  const oneDriveUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Documents/${fileName}.docx:/content`;

  const response = await fetch(oneDriveUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    },
    body: Buffer.from(wordContent, 'base64') // Convert base64 to binary
  });

  if (response.ok) {
    console.log("File uploaded successfully!");
  } else {
    console.error("Upload failed", await response.text());
  }
}

async function addEntry(tableName: string = 'LivreJournal') {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem(tableName);
    const columns = table.columns.getCount();
    await context.sync();
    table.rows.add(-1, getNewRow(columns.value), true);
    table.getDataBodyRange().load('rowCount');
    await context.sync();

    [5, 6].forEach(i => {
      const cell = table.getRange()
        .getCell(table.getDataBodyRange().rowCount - 1, i);
      console.log('cell = ', cell);
      cell.numberFormatLocal = [["hh:mm:ss"]];
    });

    await context.sync();
  });

  function getNewRow(columns: number) {
    const newRow = Array(columns).map(el => '') as any[];
    const inputs = Array.from(document.getElementsByTagName('input')).filter(input => input.dataset.index);
    console.log('inputs = ', inputs)
    if (inputs.length < 1) return;

    inputs.forEach(input => {
      const index = Number(input.dataset.index);
      let value: string | number | Date = input.value;
      if (input.type === 'number')
        value = parseFloat(value);
      else if (input.type === 'date' && input.valueAsDate)
        //@ts-ignore
        value = [String(input.valueAsDate?.getDay()).padStart(2, '0'), String(input.valueAsDate.getMonth() + 1).padStart(2, '0'), String(input.valueAsDate?.getFullYear())].join('/');
      else if (input.type === 'time' && input.valueAsDate) value = [input.valueAsDate?.getHours().toString().padStart(2, '0'), input.valueAsDate?.getMinutes().toString().padStart(2, '0'), '00'].join(':');

      newRow[index] = value;
    });

    console.log('newRow = ', newRow);
    return [newRow];

    function convertTo24HourFormat(time12h: string): string {
      const [time, modifier] = time12h.split(' ');
      let [hours, minutes] = time.split(':');

      if (hours === '12') hours = '00';

      if (modifier === 'PM') hours = String(parseInt(hours, 10) + 12);

      return `${hours}:${minutes}:00`;
    }
  }
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
  const clientId = "157dd297-447d-4592-b2d3-76b643b97132"; //the new one
  const tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
  //const redirectUri = "https://script-lab.public.cdn.office.net";
  //const redirectUri = "msal157dd297-447d-4592-b2d3-76b643b97132://auth";
  const redirectUri = "https://mbibawi.github.io/ExcelInvoicingAddIn";
  // MSAL configuration
  const msalConfig = {
    auth: {
      clientId: clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: redirectUri,
    },
    cache: {
      cacheLocation: "ExcelAddIn",
      storeAuthStateInCookie: true
    }
  };
  //@ts-ignore
  const msalInstance = new msal.PublicClientApplication(msalConfig);
  const loginRequest = { scopes: ["Files.ReadWrite"]};

  return acquireToken();

  // Function to check existing authentication context
  async function acquireToken(): Promise<string | undefined | void> {
    try {
      const account = msalInstance.getAllAccounts()[0];
      if (account) {
        return acquireTokenSilently(account);
      } else {
        return loginWithPopup();
        //return loginAndGetToken();
        //openLoginWindow()
        //return getOfficeToken()
        //return getTokenWithSSO('minabibawi@gmail.com')
        //return credentitalsToken()
      }
    } catch (error) {
      console.error("Failed to acquire token from acquireToken(): ", error);
    }
  }

  async function loginWithPopup() {
    try {
      const loginResponse = await msalInstance.loginPopup(loginRequest);
      console.log('loginResponse = ', loginResponse);

      msalInstance.setActiveAccount(loginResponse.account);

      const tokenResponse = await msalInstance.acquireTokenSilent({
        account: loginResponse.account,
        scopes: ["Files.ReadWrite"]
      });

      console.log("Token acquired from loginWithPopup: ", tokenResponse.accessToken);
      return tokenResponse.accessToken
    } catch (error) {
        console.error("Error acquiring token from loginWithPopup(): ", error);
      //@ts-ignore
      if (error instanceof InteractionRequiredAuthError) {
        // Fallback to popup if silent token acquisition fails
        const response = await msalInstance.acquireTokenPopup({
          scopes: ["Files.ReadWrite"]
        });
        console.log("Token acquired via popup:", response.accessToken);
        return response.accessToken
      }
    }
  }

  async function credentitalsToken() {
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        //clientSecret: clientSecret,
      }
    }
    //@ts-ignore
    const cca = new msal.application.ConfidentialClientApplication(msalConfig);

    const tokenRequest = {
      scopes: ["Files.ReadWrite"],
    }

    try {
      const response = await cca.acquireTokenByClientCredential(tokenRequest);
      return response.accessToken;
    } catch (error) {
      console.log('Error acquiring Token: ', error)
      return null

    }

  }

  async function getOfficeToken() {
    try {
      //@ts-ignore
      return await OfficeRuntime.auth.getAccessToken()

    } catch (error) {
      console.log("Error : ", error)

    }

  }

  async function getTokenWithSSO(email: string) {
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: redirectUri,
        navigateToLoginRequestUrl: true,
      },
      cache: {
        cacheLocation: "ExcelAddIn",
        storeAuthStateInCookie: true
      }
    };
    try {
      //@ts-ignore
      const response = await msal.PublicClientApplication(msalConfig).ssoSilent({
        scopes: ["Files.ReadWrite"],
        //scopes: ["https://graph.microsoft.com/.default"],
        loginHint: email // Forces MSAL to recognize the signed-in user
      });

      console.log("Token acquired via SSO:", response.accessToken);
      return response.accessToken;
    } catch (error) {
      console.error("SSO silent authentication failed:", error);
      return null;
    }
  }

  function openLoginWindow() {
    const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=https://graph.microsoft.com/.default`;

    // Open in a new window (only works if triggered by user action)
    const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");

    if (!authWindow) {
      console.error("Popup blocked! Please allow popups.");
    }
  }

  // Function to handle login and acquire token
  async function loginAndGetToken(): Promise<string | undefined> {
    const msalConfig = {
      auth: {
        clientId: clientId,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: redirectUri
      },

      cache: {
        cacheLocation: "ExcelInvoicing", // Specify cache location
        storeAuthStateInCookie: true  // Set this to true for IE 11
      }
    };
    //@ts-ignore
    const msalInstance = new msal.PublicClientApplication(msalConfig);
    return await acquire();
    async function acquire() {
      try {
        const response = await msalInstance.handleRedirectPromise();
        if (response !== null) {
          console.log("Login successful:", response);
          return response.accessToken;
        }
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          const tokenResponse = await msalInstance.acquireTokenSilent({
            account: accounts[0],
            scopes: ["https://graph.microsoft.com/.default"]
          });
          console.log("Token acquired silently:", tokenResponse.accessToken);
          return tokenResponse.accessToken;
        }
      } catch (error) {
        console.error("Error acquiring token:", error);
        //@ts-ignore
        if (error instanceof msal.InteractionRequiredAuthError) {
          msalInstance.acquireTokenRedirect({
            scopes: ["https://graph.microsoft.com/.default"]
          });
        }
      }
    }


    return
    try {
      const loginRequest = {
        scopes: ["Files.ReadWrite"] // OneDrive scopes
      };
      await msalInstance.loginRedirect(loginRequest);

      return handleRedirectResponse();
    } catch (error) {
      console.error("Login error:", error);
      return undefined;
    }
    // Function to handle redirect response
    async function handleRedirectResponse(): Promise<string | undefined> {
      try {
        const authResult = await msalInstance.handleRedirectPromise();
        if (authResult && authResult.accessToken) {
          console.log("Access token:", authResult.accessToken);
          return authResult.accessToken;
        }
      } catch (error) {
        console.error("Redirect handling error:", error);
      }
      return undefined;
    }
  }
  // Function to get access token silently
  async function acquireTokenSilently(account: any): Promise<string | undefined> {
    try {
      const tokenRequest = {
        account: account,
        scopes: ["Files.ReadWrite"], // OneDrive scopes
      };

      const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
      if (tokenResponse && tokenResponse.accessToken) {
        console.log("Token acquired silently :", tokenResponse.accessToken);

        return tokenResponse.accessToken;
      }
    } catch (error) {
      console.error("Token silent acquisition error:", error);
    }
  }
}


async function createWordDocumentFromTemplate(templatePath: string, newDocumentPath: string, accessToken: string) {

  if (!accessToken) return;

  const headers = {
    "Authorization": `Bearer ${accessToken}`,
    "Content-Type": "application/json",
  };
  // Fetch the template file from OneDrive
  const templateResponse = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${templatePath}:/content`,
    {
      method: 'GET',
      headers: headers
    }
  );

  const templateBlob = await templateResponse.blob();
  const templateArrayBuffer = await templateBlob.arrayBuffer();
  const templateBase64 = btoa(String.fromCharCode(...new Uint8Array(templateArrayBuffer)));

  // Create the new document with the template content
  const newDocumentResponse = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${newDocumentPath}:/content`,
    {
      method: 'PUT',
      headers: headers,
      body: JSON.stringify({
        "@microsoft.graph.conflictBehavior": "rename",
        "file": {
          "@odata.type": "#microsoft.graph.file"
        },
        "fileSystemInfo": {},
        "contentBytes": templateBase64
      })
    }
  );

  const result = await newDocumentResponse.json();
  console.log("Document created:", result);
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



