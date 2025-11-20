let userLoggedIn = false;

function updateUI() {
  const loginDiv = document.getElementById("login_screen");
  const projectDiv = document.getElementById("add_project");
  const insertBtn = document.getElementById("insert_taskpane");
  const infoMsg = document.getElementById("info_message");

  if (userLoggedIn) {
    // ✅ Logged in: show project UI
    if (loginDiv) loginDiv.style.display = "none";
    if (projectDiv) projectDiv.style.display = "block";
    if (insertBtn) insertBtn.style.display = "inline-block";
    if (infoMsg) infoMsg.textContent = "";
  } else {
    // ❌ Not logged in: show login form
    if (loginDiv) loginDiv.style.display = "block";
    if (projectDiv) projectDiv.style.display = "none";
    if (insertBtn) insertBtn.style.display = "none";
    if (infoMsg) {
      infoMsg.textContent = "First login to use 'Addin projects'.";
    }
  }
}

async function SendUserData() {
  const email = document.getElementById("email")?.value;
  const password = document.getElementById("password")?.value;
  const url = "http://127.0.0.1:8000/login/";

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });

    const data = await response.json();
    console.log("Login response:", data);

    if (response.ok) {
      userLoggedIn = true;
      await OfficeRuntime.storage.setItem("userLoggedIn", "true");
      console.log("Login successful!");
    } else {
      userLoggedIn = false;
      await OfficeRuntime.storage.setItem("userLoggedIn", "false");
      console.log("Login failed. Check your email/password.");
    }
  } catch (error) {
    console.error("Login error:", error);
    userLoggedIn = false;
    await OfficeRuntime.storage.setItem("userLoggedIn", "false");
  }

  // After login attempt, refresh UI (switch screens)
  updateUI();
}

// Optional: some Excel demo logic for the taskpane Insert button
async function insertFromTaskpane() {
  try {
    if (!userLoggedIn) {
      console.log("Please login first.");
      return;
    }

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:C1");
      range.values = [["Material", "Qty", "Price"]];
      await context.sync();
    });

    console.log("Inserted header row from taskpane.");
  } catch (err) {
    console.error("insertFromTaskpane error:", err);
  }
}

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Taskpane ready");

    // Restore login state from previous session
    try {
      const stored = await OfficeRuntime.storage.getItem("userLoggedIn");
      userLoggedIn = stored === "true";
    } catch (e) {
      userLoggedIn = false;
    }

    // Set handlers
    const loginBtn = document.getElementById("login_button");
    if (loginBtn) {
      loginBtn.onclick = SendUserData;
    }

    const insertBtn = document.getElementById("insert_taskpane");
    if (insertBtn) {
      insertBtn.onclick = insertFromTaskpane;
    }

    // Initial UI according to login state
    updateUI();
  }
});

