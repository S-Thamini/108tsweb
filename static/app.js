document.addEventListener("DOMContentLoaded", () => {
    const inputs = {
        index: document.getElementById("index"),
        data_date: document.getElementById("data_date"),
        expiry: document.getElementById("expiry"),
        template_name: document.getElementById("template_name"),
        dynamic_name: document.getElementById("dynamic_name"),
        target_sh: document.getElementById("target_sh"),
        auto_mid: document.getElementById("auto_mid"),
        mid_strike: document.getElementById("mid_strike"),
        rows: document.getElementById("rows"),
        gap: document.getElementById("gap")
    };

    // UI elements
    const form = document.getElementById("sync-form");
    const btn = document.getElementById("sync-btn");
    const btnIcon = document.getElementById("btn-icon");
    const btnText = document.getElementById("btn-text");
    const statusContainer = document.getElementById("status-container");
    const statusMessage = document.getElementById("status-message");

    // 1. Initial State / Auto-Saves
    function loadSavedData() {
        const saved = JSON.parse(localStorage.getItem("108ts_prefs") || "{}");
        
        Object.keys(inputs).forEach(key => {
            const input = inputs[key];
            if (saved[key] !== undefined) {
                if (input.type === "checkbox") {
                    input.checked = saved[key];
                } else {
                    input.value = saved[key];
                }
            }
        });

        // Apply initial disabled states based on checkboxes
        inputs.target_sh.disabled = inputs.dynamic_name.checked;
        inputs.mid_strike.disabled = inputs.auto_mid.checked;
    }

    function saveCurrentData() {
        const data = {};
        Object.keys(inputs).forEach(key => {
            const input = inputs[key];
            data[key] = (input.type === "checkbox") ? input.checked : input.value;
        });
        localStorage.setItem("108ts_prefs", JSON.stringify(data));
    }

    // Initialize Dates if NOT saved
    if (!localStorage.getItem("108ts_prefs")) {
        const today = new Date();
        inputs.data_date.value = today.toISOString().split('T')[0];
        
        const d = new Date();
        d.setDate(d.getDate() + (4 + 7 - d.getDay()) % 7);
        inputs.expiry.value = d.toISOString().split('T')[0];
    } else {
        loadSavedData();
    }

    // Attach Save Listeners to ALL inputs
    Object.values(inputs).forEach(input => {
        input.addEventListener("input", saveCurrentData);
        input.addEventListener("change", saveCurrentData);
    });

    // Special Toggle Handlers for UI
    inputs.dynamic_name.addEventListener("change", (e) => {
        inputs.target_sh.disabled = e.target.checked;
    });

    inputs.auto_mid.addEventListener("change", (e) => {
        inputs.mid_strike.disabled = e.target.checked;
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
        if (inputs.dynamic_name.checked) formData.set("target_sh", "Dynamic");
        if (inputs.auto_mid.checked) formData.set("mid_strike", "0");

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
