document.getElementById("extractBtn").addEventListener("click", async () => {
    const checkNumber = document.getElementById("checkNumber").value.trim();
    const message = document.getElementById("message");

    if (!checkNumber) {
        message.textContent = "⚠️ Enter a check number first.";
        message.style.color = 'red'
        return;
    }

    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ["libs/xlsx.full.min.js", "content.js"]
    });

    chrome.tabs.sendMessage(tab.id, { action: "extract", checkNumber }, (response) => {
        if (response?.status === "success") {
            message.textContent = "✅ Data exported successfully!";
            message.style.color = 'green'
        } else {
            message.textContent = "❌ No data found or extraction failed.";
            message.style.color = 'red'
        }
    });
});
