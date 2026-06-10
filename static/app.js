document.addEventListener("DOMContentLoaded", () => {
    // State Variables
    let uploadedFile = null;
    let pipelineRunning = false;
    let pipelineResults = null;
    let progressTimer = null;

    // Elements
    const toggleSettingsBtn = document.getElementById("toggle-settings-btn");
    const settingsPanel = document.getElementById("settings-panel");
    const openaiKeyInput = document.getElementById("openai-key-input");
    const anthropicKeyInput = document.getElementById("anthropic-key-input");
    const saveSettingsBtn = document.getElementById("save-settings-btn");
    const settingsStatus = document.getElementById("settings-status");
    const eyeToggles = document.querySelectorAll(".eye-toggle");

    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const fileInfo = document.getElementById("file-info");
    const fileName = document.getElementById("file-name");
    const removeFileBtn = document.getElementById("remove-file-btn");

    const tabJdText = document.getElementById("tab-jd-text");
    const tabJdUrl = document.getElementById("tab-jd-url");
    const jdTextPane = document.getElementById("jd-text-pane");
    const jdUrlPane = document.getElementById("jd-url-pane");
    const jdTextInput = document.getElementById("jd-text-input");
    const jdUrlInput = document.getElementById("jd-url-input");
    const scrapeBtn = document.getElementById("scrape-btn");
    const scrapedPreviewContainer = document.getElementById("scraped-preview-container");
    const scrapedPreview = document.getElementById("scraped-preview");

    const optimizeBtn = document.getElementById("optimize-btn");
    const pipelineStatus = document.getElementById("pipeline-status");
    const outputPlaceholder = document.getElementById("output-placeholder");
    const outputContainer = document.getElementById("output-container");
    
    // Timeline steps
    const stepParsing = document.getElementById("step-parsing");
    const stepTailor = document.getElementById("step-tailor");
    const stepAts = document.getElementById("step-ats");
    const stepClaude = document.getElementById("step-claude");
    const stepHumanizer = document.getElementById("step-humanizer");
    const atsBadgeContainer = document.getElementById("ats-badge-container");

    // Results elements
    const outputTabs = document.querySelectorAll(".output-tab");
    const tabPanes = document.querySelectorAll(".tab-pane");
    const activeTabTitle = document.getElementById("active-tab-title");
    const copyBtn = document.getElementById("copy-btn");
    const downloadMdBtn = document.getElementById("download-md-btn");
    const downloadDocxBtn = document.getElementById("download-docx-btn");
    const downloadPdfBtn = document.getElementById("download-pdf-btn");
    const downloadTxtBtn = document.getElementById("download-txt-btn");

    const finalResumeMarkdown = document.getElementById("final-resume-markdown");
    const coverLetterMarkdown = document.getElementById("cover-letter-markdown");
    const initialResumeMarkdown = document.getElementById("initial-resume-markdown");
    
    const scoreStroke = document.getElementById("score-stroke");
    const scoreNumber = document.getElementById("score-number");
    const scoreBadge = document.getElementById("score-badge");
    const circularChart = scoreStroke.closest(".circular-chart");
    const atsKeywordsList = document.getElementById("ats-keywords-list");
    const atsFormattingList = document.getElementById("ats-formatting-list");
    const atsSuggestionsList = document.getElementById("ats-suggestions-list");

    const auditStatusBadge = document.getElementById("audit-status-badge");
    const auditSummaryText = document.getElementById("audit-summary-text");
    const claudeFeedbackContent = document.getElementById("claude-feedback-content");

    // Load saved settings
    loadAPIKeys();

    // Toggle Settings Panel
    toggleSettingsBtn.addEventListener("click", () => {
        settingsPanel.classList.toggle("hidden");
    });

    // Eye toggle for password fields
    eyeToggles.forEach(toggle => {
        toggle.addEventListener("click", () => {
            const targetId = toggle.getAttribute("data-target");
            const input = document.getElementById(targetId);
            const icon = toggle.querySelector("i");
            if (input.type === "password") {
                input.type = "text";
                icon.classList.replace("fa-eye", "fa-eye-slash");
            } else {
                input.type = "password";
                icon.classList.replace("fa-eye-slash", "fa-eye");
            }
        });
    });

    // Save API keys
    saveSettingsBtn.addEventListener("click", () => {
        const oKey = openaiKeyInput.value.trim();
        const aKey = anthropicKeyInput.value.trim();
        
        localStorage.setItem("openai_api_key", oKey);
        localStorage.setItem("anthropic_api_key", aKey);

        showSettingsStatus("API Keys configured successfully!", "success");
        setTimeout(() => {
            settingsPanel.classList.add("hidden");
        }, 1200);
    });

    function loadAPIKeys() {
        const oKey = localStorage.getItem("openai_api_key") || "";
        const aKey = localStorage.getItem("anthropic_api_key") || "";
        openaiKeyInput.value = oKey;
        anthropicKeyInput.value = aKey;
        
        if (oKey || aKey) {
            toggleSettingsBtn.innerHTML = `<i class="fa-solid fa-gear"></i> Settings <span style="color:#10b981; margin-left:4px;">●</span>`;
        }
    }

    function showSettingsStatus(msg, type) {
        settingsStatus.textContent = msg;
        settingsStatus.className = `status-msg ${type}`;
        setTimeout(() => {
            settingsStatus.textContent = "";
            settingsStatus.className = "status-msg";
        }, 4000);
    }

    // ==========================================
    // File Upload Drag & Drop Handles
    // ==========================================
    dropZone.addEventListener("click", () => fileInput.click());

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });

    ["dragleave", "drop"].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove("dragover");
        });
    });

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        if (e.dataTransfer.files.length > 0) {
            handleFileSelection(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener("change", () => {
        if (fileInput.files.length > 0) {
            handleFileSelection(fileInput.files[0]);
        }
    });

    function handleFileSelection(file) {
        const validTypes = [".pdf", ".docx", ".txt"];
        const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();

        if (ext === ".doc") {
            alert("Legacy .doc files aren't supported. Please open your resume in Word and use 'Save As' to convert it to .docx, then re-upload.");
            return;
        }
        if (!validTypes.includes(ext)) {
            alert("Invalid file format. Please upload a PDF, DOCX, or TXT file.");
            return;
        }

        if (file.size > 5 * 1024 * 1024) {
            alert("File size exceeds 5MB limit.");
            return;
        }

        uploadedFile = file;
        fileName.textContent = file.name;
        fileInfo.classList.remove("hidden");
        dropZone.style.display = "none";
    }

    removeFileBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        uploadedFile = null;
        fileInput.value = "";
        fileInfo.classList.add("hidden");
        dropZone.style.display = "flex";
    });

    // ==========================================
    // Job Description Tab switching & Scraping
    // ==========================================
    tabJdText.addEventListener("click", () => {
        tabJdText.classList.add("active");
        tabJdUrl.classList.remove("active");
        jdTextPane.classList.remove("hidden");
        jdUrlPane.classList.add("hidden");
    });

    tabJdUrl.addEventListener("click", () => {
        tabJdUrl.classList.add("active");
        tabJdText.classList.remove("active");
        jdUrlPane.classList.remove("hidden");
        jdTextPane.classList.add("hidden");
    });

    scrapeBtn.addEventListener("click", async () => {
        const url = jdUrlInput.value.trim();
        if (!url) {
            alert("Please enter a valid job URL.");
            return;
        }

        scrapeBtn.disabled = true;
        scrapeBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Scraping...`;
        scrapedPreviewContainer.classList.add("hidden");

        try {
            const formData = new FormData();
            formData.append("url", url);

            const res = await fetch("/api/scrape-jd", {
                method: "POST",
                body: formData
            });

            if (!res.ok) {
                const errData = await res.json();
                throw new Error(errData.detail || "Failed to scrape job link.");
            }

            const data = await res.json();
            scrapedPreview.textContent = data.text;
            scrapedPreviewContainer.classList.remove("hidden");
        } catch (err) {
            alert(err.message);
        } finally {
            scrapeBtn.disabled = false;
            scrapeBtn.innerHTML = `<i class="fa-solid fa-globe"></i> Scrape`;
        }
    });

    // ==========================================
    // Pipeline Launcher & Progress Sim
    // ==========================================
    optimizeBtn.addEventListener("click", async () => {
        if (pipelineRunning) return;

        // Validation
        if (!uploadedFile) {
            alert("Please upload an existing resume first.");
            return;
        }

        const isTextMode = tabJdText.classList.contains("active");
        const jdText = jdTextInput.value.trim();
        const jdUrl = jdUrlInput.value.trim();

        if (isTextMode && !jdText) {
            alert("Please paste the target job description text.");
            return;
        }
        if (!isTextMode && !jdUrl) {
            alert("Please provide the job description link.");
            return;
        }

        const oKey = localStorage.getItem("openai_api_key") || "";
        const aKey = localStorage.getItem("anthropic_api_key") || "";
        
        // Show loading panels
        pipelineRunning = true;
        optimizeBtn.disabled = true;
        optimizeBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Processing Pipeline...`;
        
        outputPlaceholder.classList.add("hidden");
        outputContainer.classList.add("hidden");
        pipelineStatus.classList.remove("hidden");

        // Start simulated timeline steps (for better UX since total wait is ~45s)
        resetTimeline();
        startTimelineSimulation();

        try {
            const formData = new FormData();
            formData.append("resume", uploadedFile);
            if (isTextMode) {
                formData.append("jd_text", jdText);
            } else {
                formData.append("jd_url", jdUrl);
            }
            formData.append("openai_key", oKey);
            formData.append("anthropic_key", aKey);

            const res = await fetch("/api/optimize", {
                method: "POST",
                body: formData
            });

            if (!res.ok) {
                const errData = await res.json();
                throw new Error(errData.detail || "An error occurred in the optimization pipeline.");
            }

            const results = await res.json();
            pipelineResults = results;
            
            // Finish timeline instantly
            completeAllTimelineSteps(results);
            
            // Render outputs
            renderPipelineResults(results);

            // Display Results
            pipelineStatus.classList.add("hidden");
            outputContainer.classList.remove("hidden");
        } catch (err) {
            alert(err.message);
            pipelineStatus.classList.add("hidden");
            outputPlaceholder.classList.remove("hidden");
            
            // If the error is key-related, auto expand settings
            if (err.message.includes("key") || err.message.includes("API key")) {
                settingsPanel.classList.remove("hidden");
            }
        } finally {
            clearInterval(progressTimer);
            pipelineRunning = false;
            optimizeBtn.disabled = false;
            optimizeBtn.innerHTML = `<i class="fa-solid fa-rocket"></i> Launch Optimization Pipeline`;
        }
    });

    // Progress timeline mechanics
    function resetTimeline() {
        const steps = [stepParsing, stepTailor, stepAts, stepClaude, stepHumanizer];
        steps.forEach(step => {
            step.className = "timeline-item";
        });
        atsBadgeContainer.innerHTML = "";
    }

    function startTimelineSimulation() {
        let elapsed = 0;
        stepParsing.classList.add("active");

        progressTimer = setInterval(() => {
            elapsed += 1;

            if (elapsed === 3) {
                stepParsing.classList.remove("active");
                stepParsing.classList.add("completed");
                stepTailor.classList.add("active");
            } else if (elapsed === 15) {
                stepTailor.classList.remove("active");
                stepTailor.classList.add("completed");
                stepAts.classList.add("active");
                
                // Mock visual check feedback
                atsBadgeContainer.innerHTML = `<span class="attempt-badge low">Check #1: 72%</span>`;
            } else if (elapsed === 22) {
                atsBadgeContainer.innerHTML += `<span class="attempt-badge low">Check #2: 81%</span>`;
            } else if (elapsed === 28) {
                atsBadgeContainer.innerHTML += `<span class="attempt-badge pass">Check #3: 88%</span>`;
                stepAts.classList.remove("active");
                stepAts.classList.add("completed");
                stepClaude.classList.add("active");
            } else if (elapsed === 45) {
                stepClaude.classList.remove("active");
                stepClaude.classList.add("completed");
                stepHumanizer.classList.add("active");
            }
        }, 1000);
    }

    function completeAllTimelineSteps(results) {
        clearInterval(progressTimer);
        const steps = [stepParsing, stepTailor, stepAts, stepClaude, stepHumanizer];
        steps.forEach(step => {
            step.classList.remove("active");
            step.classList.add("completed");
        });

        // Load real ATS history badges
        atsBadgeContainer.innerHTML = "";
        if (results.ats_history) {
            results.ats_history.forEach(item => {
                const isPass = item.score >= 85;
                const statusClass = isPass ? "pass" : "low";
                atsBadgeContainer.innerHTML += `<span class="attempt-badge ${statusClass}">Attempt #${item.attempt}: ${item.score}%</span>`;
            });
        }
    }

    // ==========================================
    // Render Results & Tabs Management
    // ==========================================
    function renderPipelineResults(data) {
        // Render markdown content (overflow-safe via .markdown-preview wrap rules)
        finalResumeMarkdown.innerHTML = marked.parse(data.final_resume || "");
        coverLetterMarkdown.innerHTML = marked.parse(data.cover_letter || "");
        initialResumeMarkdown.innerHTML = marked.parse(data.initial_tailored_resume || "");

        // Render ATS score
        const lastAts = data.ats_history[data.ats_history.length - 1];
        const finalScore = lastAts ? lastAts.score : 0;
        
        scoreNumber.textContent = `${finalScore}%`;
        
        // Circular progress svg calculation
        // stroke-dasharray="score, 100" maps score directly to percentage circle fill
        scoreStroke.setAttribute("stroke-dasharray", `${finalScore}, 100`);

        // Set style class and badge text based on score
        circularChart.className.baseVal = "circular-chart"; // reset
        if (finalScore < 70) {
            circularChart.classList.add("low-score");
            scoreBadge.textContent = "Poor Fit";
            scoreBadge.className = "score-badge low";
        } else if (finalScore < 85) {
            circularChart.classList.add("mid-score");
            scoreBadge.textContent = "Average Fit";
            scoreBadge.className = "score-badge mid";
        } else {
            circularChart.classList.add("high-score");
            scoreBadge.textContent = "Strong Fit";
            scoreBadge.className = "score-badge high";
        }

        // Render keywords and suggestions
        const feedback = lastAts ? lastAts.feedback : {};
        renderList(atsKeywordsList, feedback.missing_keywords || ["None identified"]);
        renderList(atsFormattingList, feedback.formatting_issues || ["No significant formatting issues detected"]);
        renderList(atsSuggestionsList, feedback.suggestions || ["All criteria successfully matched"]);

        // Render Claude review audit
        if (data.claude_updated) {
            auditStatusBadge.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Enhanced by Claude`;
            auditStatusBadge.className = "audit-status-badge low"; // reuse amber styling
            auditSummaryText.textContent = "Claude updated specific action phrases and structured bullet points to read more professional.";
        } else {
            auditStatusBadge.innerHTML = `<i class="fa-solid fa-circle-check"></i> Passed Audit`;
            auditStatusBadge.className = "audit-status-badge";
            auditSummaryText.textContent = "Claude reviewed the tailored resume and found no substantial updates needed.";
        }
        claudeFeedbackContent.textContent = data.claude_feedback;
        
        // Default to showing Final Resume tab on reload
        switchTab("tab-final-resume");
    }

    function renderList(container, list) {
        container.innerHTML = "";
        if (list.length === 0) {
            container.innerHTML = "<li>None</li>";
            return;
        }
        list.forEach(item => {
            const li = document.createElement("li");
            li.textContent = item;
            container.appendChild(li);
        });
    }

    // Tabs navigation inside results card
    outputTabs.forEach(tab => {
        tab.addEventListener("click", () => {
            const targetPaneId = tab.getAttribute("data-tab");
            switchTab(targetPaneId);
        });
    });

    function switchTab(paneId) {
        outputTabs.forEach(tab => {
            if (tab.getAttribute("data-tab") === paneId) {
                tab.classList.add("active");
            } else {
                tab.classList.remove("active");
            }
        });

        tabPanes.forEach(pane => {
            if (pane.id === paneId) {
                pane.classList.add("active-pane");
            } else {
                pane.classList.remove("active-pane");
            }
        });

        // Set action headers
        const tabMap = {
            "tab-final-resume": "Optimized Resume",
            "tab-cover-letter": "Tailored Cover Letter",
            "tab-ats-report": "ATS Score Breakdown",
            "tab-claude-audit": "Claude Recruiter Audit",
            "tab-initial-resume": "Initial Draft Resume"
        };
        activeTabTitle.textContent = tabMap[paneId] || "Results";
    }

    // ==========================================
    // Export Operations
    // ==========================================
    
    // Copy to clipboard
    copyBtn.addEventListener("click", () => {
        if (!pipelineResults) return;
        
        const activePane = document.querySelector(".tab-pane.active-pane");
        let textToCopy = "";

        if (activePane.id === "tab-final-resume") {
            textToCopy = pipelineResults.final_resume;
        } else if (activePane.id === "tab-cover-letter") {
            textToCopy = pipelineResults.cover_letter;
        } else if (activePane.id === "tab-initial-resume") {
            textToCopy = pipelineResults.initial_tailored_resume;
        } else if (activePane.id === "tab-ats-report") {
            textToCopy = JSON.stringify(pipelineResults.ats_history, null, 2);
        } else if (activePane.id === "tab-claude-audit") {
            textToCopy = pipelineResults.claude_feedback;
        }

        navigator.clipboard.writeText(textToCopy)
            .then(() => {
                const prevHtml = copyBtn.innerHTML;
                copyBtn.innerHTML = `<i class="fa-solid fa-check"></i> Copied!`;
                setTimeout(() => copyBtn.innerHTML = prevHtml, 2000);
            })
            .catch(err => {
                alert("Failed to copy text: " + err);
            });
    });

    // Download as Markdown file
    downloadMdBtn.addEventListener("click", () => {
        if (!pipelineResults) return;
        
        const activePane = document.querySelector(".tab-pane.active-pane");
        let content = "";
        let filename = "";

        if (activePane.id === "tab-final-resume") {
            content = pipelineResults.final_resume;
            filename = "Optimized_Resume.md";
        } else if (activePane.id === "tab-cover-letter") {
            content = pipelineResults.cover_letter;
            filename = "Cover_Letter.md";
        } else if (activePane.id === "tab-initial-resume") {
            content = pipelineResults.initial_tailored_resume;
            filename = "Initial_Draft_Resume.md";
        } else {
            alert("Markdown download is only available for text documents (Resumes or Cover Letter).");
            return;
        }

        const blob = new Blob([content], { type: "text/markdown;charset=utf-8;" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.setAttribute("download", filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    // ==========================================
    // Server-side downloads (Word .docx, PDF, Plain text)
    // The server strips AI/tool fingerprints (metadata, smart quotes, emoji)
    // and applies clean professional formatting.
    // ==========================================
    function getActivePaneContent() {
        if (!pipelineResults) return null;
        const activePane = document.querySelector(".tab-pane.active-pane");
        if (activePane.id === "tab-final-resume") {
            return { content: pipelineResults.final_resume, name: "Optimized_Resume" };
        } else if (activePane.id === "tab-cover-letter") {
            return { content: pipelineResults.cover_letter, name: "Cover_Letter" };
        } else if (activePane.id === "tab-initial-resume") {
            return { content: pipelineResults.initial_tailored_resume, name: "Initial_Draft_Resume" };
        }
        return null;
    }

    async function downloadFromServer(endpoint, ext) {
        const item = getActivePaneContent();
        if (!item) {
            alert(`This download is only available for Resume and Cover Letter tabs.`);
            return;
        }
        try {
            const res = await fetch(endpoint, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ content: item.content, filename: item.name })
            });
            if (!res.ok) {
                let detail = `Download failed (${res.status})`;
                try { detail = (await res.json()).detail || detail; } catch (_) {}
                throw new Error(detail);
            }
            const blob = await res.blob();
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.setAttribute("download", `${item.name}.${ext}`);
            document.body.appendChild(link);
            link.click();
            URL.revokeObjectURL(link.href);
            document.body.removeChild(link);
        } catch (err) {
            alert(err.message);
        }
    }

    downloadDocxBtn.addEventListener("click", () => downloadFromServer("/api/download/docx", "docx"));
    downloadPdfBtn.addEventListener("click", () => downloadFromServer("/api/download/pdf", "pdf"));
    downloadTxtBtn.addEventListener("click", () => downloadFromServer("/api/download/txt", "txt"));
});
