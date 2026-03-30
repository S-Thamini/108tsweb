document.addEventListener("DOMContentLoaded", () => {
    // Inputs
    const dynamicNameCheck = document.getElementById("dynamic_name");
    const targetShInput = document.getElementById("target_sh");
    
    const autoMidCheck = document.getElementById("auto_mid");
    const midStrikeInput = document.getElementById("mid_strike");

    const dataDateInput = document.getElementById("data_date");
    const expiryInput = document.getElementById("expiry");

    // UI elements
    const form = document.getElementById("sync-form");
    const btn = document.getElementById("sync-btn");
    const btnIcon = document.getElementById("btn-icon");
    const btnText = document.getElementById("btn-text");
    const statusContainer = document.getElementById("status-container");
    const statusMessage = document.getElementById("status-message");

    // Initialize Dates
    const today = new Date();
    dataDateInput.value = today.toISOString().split('T')[0];
    
    // Find next Thursday
    const d = new Date();
    d.setDate(d.getDate() + (4 + 7 - d.getDay()) % 7);
    expiryInput.value = d.toISOString().split('T')[0];

    // Toggle Handlers
    dynamicNameCheck.addEventListener("change", (e) => {
        targetShInput.disabled = e.target.checked;
    });

    autoMidCheck.addEventListener("change", (e) => {
        midStrikeInput.disabled = e.target.checked;
    });

    // Form Submit
    form.addEventListener("submit", async (e) => {
        e.preventDefault();
        
        // UI Loading State
        btn.disabled = true;
        btnIcon.className = "ri-loader-4-line spin";
        btnText.innerText = "Processing Data...";
        statusContainer.classList.add("hidden");

        const formData = new FormData(form);
        
        // Ensure disabled values are sent (FormData ignores disabled inputs)
        if (dynamicNameCheck.checked) formData.set("target_sh", "Dynamic");
        if (autoMidCheck.checked) formData.set("mid_strike", "0");

        try {
            const response = await fetch("/api/sync", {
                method: "POST",
                body: formData
            });

            const result = await response.json();

            if (response.ok) {
                showStatus("Success! " + result.message, "success");
                
                // Trigger Download automatically
                const downloadLink = document.createElement("a");
                downloadLink.href = result.download_url;
                downloadLink.setAttribute("download", "");
                document.body.appendChild(downloadLink);
                downloadLink.click();
                document.body.removeChild(downloadLink);
                
                showStatus("File Downloaded Successfully. You can open it in Excel.", "success");
            } else {
                throw new Error(result.detail || "An error occurred during sync.");
            }
        } catch (error) {
            showStatus(error.message, "error");
        } finally {
            // Restore UI
            btn.disabled = false;
            btnIcon.className = "ri-loop-right-line";
            btnText.innerText = "Synchronize & Download";
        }
    });

    function showStatus(message, type) {
        statusContainer.className = `status-container ${type === 'error' ? 'error' : ''}`;
        statusContainer.querySelector('.status-icon i').className = type === 'error' ? 'ri-error-warning-line' : 'ri-check-line';
        statusMessage.innerText = message;
    }
});
