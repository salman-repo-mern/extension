chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "extract") {
        const checkNumber = request.checkNumber;
        const rows = document.querySelectorAll("table tr");
        let results = [];

        rows.forEach((row) => {
            const cells = row.querySelectorAll("td");
            if (cells.length > 0 && cells[0].innerText.trim() === checkNumber) {
                const paymentId = cells[1]?.innerText.trim() || "N/A";
                const patientCount = cells[2]?.innerText.trim() || "N/A";
                results.push({ checkNumber, paymentId, patientCount });
            }
        });

        if (results.length > 0) {
            const ws = XLSX.utils.json_to_sheet(results);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "ECW_Data");

            const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
            const blob = new Blob([wbout], { type: "application/octet-stream" });
            const url = URL.createObjectURL(blob);

            const a = document.createElement("a");
            a.href = url;
            a.download = "ecw_data.xlsx";
            a.click();

            sendResponse({ status: "success" });
        } else {
            sendResponse({ status: "not_found" });
        }
    }
    return true;
});
