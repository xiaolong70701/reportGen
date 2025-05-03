let googleClientId = "";
let googleDeveloperKey = "";
let pickerApiLoaded = false;
let oauthToken = null;
let tokenClient = null;
let currentFileType = null;

// 載入 config 並初始化 Token Client
fetch("/config")
  .then((res) => res.json())
  .then((config) => {
    console.log("[✅] config loaded", config);
    googleDeveloperKey = config.GOOGLE_DEVELOPER_KEY;
    googleClientId = config.GOOGLE_CLIENT_ID;

    // 初始化 Token Client
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: googleClientId,
      scope: "https://www.googleapis.com/auth/drive.readonly",
      callback: (tokenResponse) => {
        oauthToken = tokenResponse.access_token;
        console.log("[🔐] Token acquired:", oauthToken);

        // 傳送 token 給後端
        fetch("/set_token", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token: oauthToken }),
        });

        // 標記登入
        sessionStorage.setItem("drive_logged_in", "true");

        // 有待選檔案類型則開啟 Picker
        if (currentFileType) createPicker(currentFileType);
      },
    });

    // 載入 Picker API
    if (window.gapi && gapi.load) {
      gapi.load("picker", onPickerApiLoad);
    } else {
      console.error("❌ 無法載入 gapi，請確認 api.js 是否正確引入");
    }
  });

// Picker 載入完成
function onPickerApiLoad() {
  pickerApiLoaded = true;
  console.log("[✅] Picker API 載入完成");
}

// 啟動 Google Picker
function startPicker(fileType) {
  if (!googleClientId || !googleDeveloperKey) {
    alert("Google API 設定尚未正確載入");
    return;
  }

  if (!pickerApiLoaded) {
    alert("Google Picker 尚未載入完成");
    return;
  }

  currentFileType = fileType;

  // 關閉 Modal（由 index.html 控制，如果要強制也可寫這行）
  const modalEl = document.getElementById("uploadSourceModal");
  if (modalEl) bootstrap.Modal.getInstance(modalEl)?.hide();

  // 初次授權（且未登入過）才提示授權畫面
  if (!oauthToken && !sessionStorage.getItem("drive_logged_in")) {
    tokenClient.requestAccessToken({ prompt: "" });
  } else if (!oauthToken) {
    tokenClient.requestAccessToken();
  } else {
    createPicker(fileType);
  }
}

// 建立 Picker
function createPicker(fileType) {
  if (!(pickerApiLoaded && oauthToken)) return;

  let view;
  if (fileType === "docx") {
    view = new google.picker.DocsView().setMimeTypes(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
  } else if (fileType === "csv") {
    view = new google.picker.DocsView().setMimeTypes("text/csv");
  } else if (fileType === "json") {
    view = new google.picker.DocsView().setMimeTypes("application/json");
  } else {
    view = new google.picker.View(google.picker.ViewId.DOCS);
  }

  const picker = new google.picker.PickerBuilder()
    .enableFeature(google.picker.Feature.NAV_HIDDEN)
    .setOAuthToken(oauthToken)
    .setDeveloperKey(googleDeveloperKey)
    .addView(view)
    .setCallback(pickerCallback)
    .build();

  picker.setVisible(true);
}

// Picker 選取後回傳結果
function pickerCallback(data) {
  if (data.action === google.picker.Action.PICKED) {
    const file = data.docs[0];
    console.log("[📄] 選取的檔案", file);

    const fileId = file.id;
    const fileType = file.mimeType;

    fetch(`/import_drive_file?file_id=${fileId}&type=${fileType}`)
      .then((res) => res.json())
      .then((result) => {
        if (result.success) {
          alert(`✅ 成功匯入檔案：${result.filename}`);

          const mapping = {
            docx: { nameEl: "#docx_file_name", inputEl: "#docx_file_real" },
            csv: { nameEl: "#csv_file_name", inputEl: "#csv_file_real" },
            json: { nameEl: "#settings_file_name", inputEl: "#settings_file_real" },
          };

          const target = mapping[currentFileType];
          if (target) {
            const nameNode = document.querySelector(target.nameEl);
            const inputNode = document.querySelector(target.inputEl);

            if (nameNode) nameNode.textContent = result.filename;
            if (inputNode) inputNode.removeAttribute("required");
          }
        } else {
          alert(`❌ 匯入失敗：${result.error}`);
        }
      })
      .catch((err) => {
        alert("❌ 與後端通訊錯誤，請查看 console");
        console.error(err);
      });
  }
}
