document.addEventListener('DOMContentLoaded', () => {
    const spidaFileInput = document.getElementById('spidaFile');
    const katapultFileInput = document.getElementById('katapultFile');
    const spidaDropZone = document.getElementById('spidaDropZone');
    const katapultDropZone = document.getElementById('katapultDropZone');
    const processButton = document.getElementById('processButton');
    const statusArea = document.getElementById('statusArea');

    let spidaJsonData = null;
    let katapultJsonData = null;

    // Disable process button initially
    processButton.disabled = true;

    // Add disabled button styles
    processButton.style.opacity = '0.6';
    processButton.style.cursor = 'not-allowed';

    // Function to update button state
    function updateButtonState() {
        const isEnabled = spidaJsonData && katapultJsonData;
        processButton.disabled = !isEnabled;
        processButton.style.opacity = isEnabled ? '1' : '0.6';
        processButton.style.cursor = isEnabled ? 'pointer' : 'not-allowed';
    }

    // --- Error Handling and Validation Functions ---

    /**
     * Validates the structure of SPIDAcalc JSON data.
     * @param {Object} data The SPIDAcalc JSON data to validate.
     * @returns {Object} { isValid: boolean, errors: string[] }
     */
    function validateSpidaJson(data) {
        const errors = [];
        
        // Check basic structure
        if (!data || typeof data !== 'object') {
            errors.push('Invalid SPIDAcalc JSON format: root must be an object');
            return { isValid: false, errors };
        }

        // Check for leads array
        if (!Array.isArray(data.leads)) {
            errors.push('Missing or invalid leads array in SPIDAcalc JSON');
            return { isValid: false, errors };
        }

        // Check for locations in first lead
        if (!Array.isArray(data.leads[0]?.locations)) {
            errors.push('Missing or invalid locations array in SPIDAcalc JSON');
            return { isValid: false, errors };
        }

        // Validate each location has required fields
        data.leads[0].locations.forEach((location, idx) => {
            if (!location.label) {
                errors.push(`Location at index ${idx} missing required field: label`);
            }
            if (!Array.isArray(location.designs)) {
                errors.push(`Location "${location.label || idx}" missing required field: designs array`);
            }
        });

        return {
            isValid: errors.length === 0,
            errors
        };
    }

    /**
     * Validates the structure of Katapult JSON data.
     * @param {Object} data The Katapult JSON data to validate.
     * @returns {Object} { isValid: boolean, errors: string[], warnings: string[], usableNodeCount: number }
     */
    function validateKatapultJson(data) {
        const errors = [];
        const warnings = [];
        let usableNodeCount = 0;
        
        // Check basic structure
        if (!data || typeof data !== 'object') {
            errors.push('Invalid Katapult JSON format: root must be an object');
            return { isValid: false, errors, warnings, usableNodeCount };
        }

        // Check for nodes object
        if (!data.nodes || typeof data.nodes !== 'object') {
            errors.push('Missing or invalid nodes object in Katapult JSON');
            return { isValid: false, errors, warnings, usableNodeCount };
        }

        // Validate nodes and count usable ones (with pole numbers)
        Object.entries(data.nodes).forEach(([nodeId, node]) => {
            const poleNumber = safeAccess(node, 'attributes.PoleNumber.assessment') || 
                             safeAccess(node, 'attributes.PoleNumber.-Imported');
            if (poleNumber) {
                usableNodeCount++;
            } else {
                // This is a warning, not an error
                warnings.push(`Node ${nodeId} has no PoleNumber (will be skipped during processing)`);
            }
        });

        // If no usable nodes were found, that's an error
        if (usableNodeCount === 0) {
            errors.push('No usable nodes with pole numbers found in Katapult data');
        }

        return {
            isValid: errors.length === 0,
            errors,
            warnings,
            usableNodeCount
        };
    }

    /**
     * Updates the status area with a message and optional error state.
     * Supports HTML content for better formatting.
     * @param {string} message The message to display
     * @param {boolean} isError Whether this is an error message
     * @param {boolean} isHtml Whether the message contains HTML
     */
    function updateStatus(message, isError = false, isHtml = false) {
        statusArea.style.color = isError ? 'red' : 'green';
        if (isHtml) {
            statusArea.innerHTML = message;
        } else {
            statusArea.innerHTML = `<p>${message}</p>`;
        }
    }

    /**
     * Shows a loading spinner in the status area with a message.
     * @param {string} message The message to show with the spinner
     */
    function showLoading(message) {
        const spinnerHtml = `
            <div class="loading-spinner">
                <div class="spinner"></div>
                <p>${message}</p>
            </div>
        `;
        updateStatus(spinnerHtml, false, true);
    }

    /**
     * Validates file size and type before processing.
     * @param {File} file The file to validate
     * @param {string} type 'spida' or 'katapult'
     * @returns {boolean} Whether the file is valid
     */
    function validateFile(file, type) {
        // Check file size (max 10MB)
        const maxSize = 10 * 1024 * 1024; // 10MB in bytes
        if (file.size > maxSize) {
            updateStatus(`Error: ${type} file size exceeds 10MB limit.`, true);
            return false;
        }

        // Check file type
        if (!file.name.toLowerCase().endsWith('.json')) {
            updateStatus(`Error: ${type} file must be a JSON file.`, true);
            return false;
        }

        return true;
    }

    // --- Helper Functions ---

    /**
     * Safely access nested properties of an object.
     * @param {Object} obj The object to access.
     * @param {String} path The path to the property (e.g., 'a.b.c').
     * @param {*} defaultValue The value to return if the path is not found.
     * @returns {*} The property value or the default value.
     */
    function safeAccess(obj, path, defaultValue = null) {
        if (!obj || typeof path !== 'string') {
            return defaultValue;
        }
        const keys = path.split('.');
        let current = obj;
        for (const key of keys) {
            if (current && typeof current === 'object' && key in current) {
                current = current[key];
            } else {
                return defaultValue;
            }
        }
        return current;
    }

    /**
     * Converts a numeric value (meters or decimal feet) to "X'-Y\"" format.
     * @param {number} value The numeric value.
     * @param {string} unit 'm' for meters, 'ft' for decimal feet.
     * @returns {string} Formatted height string or "NA" if input is invalid.
     */
    function convertToFeetInches(value, unit) {
        if (value === null || value === undefined || isNaN(parseFloat(value))) {
            return "NA";
        }
        let feetDecimal;
        if (unit === 'm') {
            feetDecimal = parseFloat(value) * 3.28084;
        } else if (unit === 'ft') {
            feetDecimal = parseFloat(value);
        } else {
            return "NA"; // Invalid unit
        }

        const feet = Math.floor(feetDecimal);
        const inches = Math.round((feetDecimal - feet) * 12);
        if (inches === 12) { // Handle rounding up to the next foot
            return `${feet + 1}'-0"`;
        }
        return `${feet}'-${inches}"`;
    }

    /**
     * Standardizes pole identifiers.
     * Example: "1-PL12345" -> "PL12345"
     * @param {string} poleID The original pole ID.
     * @param {string} source 'spida' or 'katapult' (can be used for source-specific rules).
     * @returns {string} Canonical pole ID or original if no specific rule applies.
     */
    function getCanonicalPoleID(poleID, source) {
        if (!poleID || typeof poleID !== 'string') return "UNKNOWN_POLE";
        // Example rule: SPIDA IDs might have a prefix like "1-" or similar.
        if (source === 'spida' && poleID.includes('-')) {
            const parts = poleID.split('-');
            if (parts.length > 1 && parts[1].trim() !== "") {
                return parts.slice(1).join('-').trim(); // Takes everything after the first hyphen
            }
        }
        // Katapult might use "PoleNumber.assessment" directly if it's already clean.
        // Add more specific rules as needed based on actual data patterns.
        return poleID.trim();
    }

    // Enhanced formatPercentage function to match target
    function formatPercentage(value) {
        const num = parseFloat(value);
        if (isNaN(num)) return "NA";
        return `${num.toFixed(2)}%`;
    }

    /**
     * Formats a count into "NO" or "YES (count)" string.
     * @param {number} count The count.
     * @returns {string} Formatted string.
     */
    function formatYesNoCount(count) {
        const num = parseInt(count, 10);
        if (isNaN(num) || num <= 0) return "NO";
        return `YES (${num})`;
    }

    // --- Drag and Drop Handlers ---

    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.add('dragover');
    }

    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.remove('dragover');
    }

    function handleDrop(e, type) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.remove('dragover');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            handleFileUpload(file, type);
            // Update the file input for consistency
            if (type === 'spida') {
                spidaFileInput.files = files;
            } else {
                katapultFileInput.files = files;
            }
        }
    }

    // Add drag and drop event listeners
    [spidaDropZone, katapultDropZone].forEach(zone => {
        zone.addEventListener('dragover', handleDragOver);
        zone.addEventListener('dragleave', handleDragLeave);
        zone.addEventListener('drop', (e) => {
            handleDrop(e, zone === spidaDropZone ? 'spida' : 'katapult');
        });
    });

    // --- Event Listeners ---
    spidaFileInput.addEventListener('change', (event) => {
        handleFileUpload(event.target.files[0], 'spida');
    });

    katapultFileInput.addEventListener('change', (event) => {
        handleFileUpload(event.target.files[0], 'katapult');
    });

    processButton.addEventListener('click', () => {
        if (!spidaJsonData || !katapultJsonData) {
            updateStatus(`
                <div class="error-message">
                    <strong>Error:</strong> Please upload both SPIDAcalc and Katapult JSON files.
                </div>
            `, true, true);
            return;
        }

        showLoading('Processing files and generating Excel report...');
        
        // Use setTimeout to allow the loading spinner to render
        setTimeout(() => {
            try {
                processAndGenerateExcel(spidaJsonData, katapultJsonData);
                updateStatus(`
                    <div class="success-message">
                        <strong>Success!</strong>
                        <br>Excel file generated and downloaded.
                        <br>You can now close this window or process another set of files.
                    </div>
                `, false, true);
            } catch (error) {
                console.error("Error during processing:", error);
                updateStatus(`
                    <div class="error-message">
                        <strong>Error during processing:</strong>
                        <br>${error.message}
                        <br>Check the browser console for more details.
                    </div>
                `, true, true);
            }
        }, 100); // Small delay to ensure loading state is shown
    });

    function handleFileUpload(file, type) {
        if (!file) {
            updateStatus(`Error: No file selected for ${type}.`, true);
            if (type === 'spida') spidaJsonData = null;
            if (type === 'katapult') katapultJsonData = null;
            return;
        }

        // Validate file size and type
        if (!validateFile(file, type)) {
            if (type === 'spida') spidaJsonData = null;
            if (type === 'katapult') katapultJsonData = null;
            return;
        }

        // Show loading state
        showLoading(`Reading ${type} file...`);

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const jsonData = JSON.parse(e.target.result);
                
                // Validate JSON structure
                const validation = type === 'spida' ? validateSpidaJson(jsonData) : validateKatapultJson(jsonData);
                
                if (!validation.isValid) {
                    const errorList = validation.errors.map(err => `<li>${err}</li>`).join('');
                    updateStatus(`
                        <div class="error-message">
                            <strong>Invalid ${type} JSON structure:</strong>
                            <ul>${errorList}</ul>
                        </div>
                    `, true, true);
                    if (type === 'spida') spidaJsonData = null;
                    if (type === 'katapult') katapultJsonData = null;
                    return;
                }

                // Store the validated data
                if (type === 'spida') {
                    spidaJsonData = jsonData;
                    updateStatus(`
                        <div class="success-message">
                            SPIDAcalc JSON loaded successfully.
                            <br>Found ${jsonData.leads[0].locations.length} poles.
                        </div>
                    `, false, true);
                } else if (type === 'katapult') {
                    katapultJsonData = jsonData;
                    
                    // Include warnings if any
                    let warningHtml = '';
                    if (validation.warnings && validation.warnings.length > 0) {
                        const truncatedWarnings = validation.warnings.slice(0, 5); // Show only first 5
                        const remainingCount = validation.warnings.length - 5;
                        
                        warningHtml = `
                            <div class="warning-message" style="margin-top: 10px;">
                                <strong>Warnings:</strong>
                                <p>${truncatedWarnings.length} nodes without pole numbers found (these will be skipped)</p>
                                ${remainingCount > 0 ? `<p>...and ${remainingCount} more similar warnings</p>` : ''}
                            </div>
                        `;
                    }
                    
                    updateStatus(`
                        <div class="success-message">
                            Katapult JSON loaded successfully.
                            <br>Found ${validation.usableNodeCount} usable nodes with pole numbers.
                        </div>
                        ${warningHtml}
                    `, false, true);
                }

                // Update process button state
                updateButtonState();
            } catch (error) {
                console.error(`Error parsing ${type} JSON:`, error);
                updateStatus(`
                    <div class="error-message">
                        <strong>Error parsing ${type} JSON:</strong>
                        <br>${error.message}
                    </div>
                `, true, true);
                if (type === 'spida') spidaJsonData = null;
                if (type === 'katapult') katapultJsonData = null;
            }
        };

        reader.onerror = () => {
            updateStatus(`
                <div class="error-message">
                    <strong>Error reading ${type} file:</strong>
                    <br>${reader.error.message}
                </div>
            `, true, true);
            if (type === 'spida') spidaJsonData = null;
            if (type === 'katapult') katapultJsonData = null;
        };

        reader.readAsText(file);
    }

    function updateStatus(message, isError = false) {
        statusArea.innerHTML = `<p>${message}</p>`;
        statusArea.style.color = isError ? 'red' : 'green';
    }

    // --- Core Processing and Excel Generation Logic ---
    function processAndGenerateExcel(spidaData, katapultData) {
        // Initialize report data array with header rows
        const reportData = [
            // Row 1 (Main Headers)
            [
                "Operation Number",
                "Attachment Action:\n( I )nstalling\n( R )emoving\n( E )xisting",
                "Pole Owner",
                "Pole #",
                "Pole Structure",
                "Proposed Riser (Yes/No) &",
                "Proposed Guy (Yes/No) &",
                "PLA (%) with proposed attachment",
                "Construction Grade of Analysis",
                "Existing Mid-Span Data", null, // Merge J1:K1
                "Make Ready Data", null, null, null // Merge L1:O1
            ],
            // Row 2 (Sub-Headers)
            [
                null, null, null, null, null, null, null, null, null,
                "Height Lowest Com",
                "Height Lowest CPS Electrical",
                "Attachment Height", null, null, // Merge L2:N2
                "Mid-Span\n(same span as existing)"
            ],
            // Row 3 (Lowest-Level Sub-Headers)
            [
                null, null, null, null, null, null, null, null, null, null, null,
                "Attacher Description",
                "Existing",
                "Proposed",
                "Proposed"
            ]
        ];

        // Track merges for header rows
        const merges = [
            { s: { r: 0, c: 9 }, e: { r: 0, c: 10 } }, // J1:K1
            { s: { r: 0, c: 11 }, e: { r: 0, c: 14 } }, // L1:O1
            { s: { r: 1, c: 11 }, e: { r: 1, c: 13 } }  // L2:N2
        ];

        // Create Katapult pole lookup map
        const katapultPoleMap = new Map();
        Object.entries(katapultData.nodes || {}).forEach(([nodeId, node]) => {
            const poleNumber = safeAccess(node, 'attributes.PoleNumber.assessment') || // Preferred
                             safeAccess(node, 'attributes.PoleNumber.-Imported') || // Fallback
                             safeAccess(node, 'attributes.pole_tag.tagtext');     // Alternative
            if (poleNumber) {
                // Ensure poleNumber is a string before calling getCanonicalPoleID
                const poleNumStr = typeof poleNumber === 'string' ? poleNumber : safeAccess(poleNumber, 'tagtext', String(poleNumber));
                const canonicalID = getCanonicalPoleID(poleNumStr, 'katapult');
                if (canonicalID) {
                    katapultPoleMap.set(canonicalID, { ...node, katapultNodeId: nodeId }); // Store nodeId
                }
            }
        });

        // Process each pole from SPIDAcalc
        let operationNumber = 0;
        const spidaPoles = safeAccess(spidaData, 'leads.0.locations', []);
        
        // TEMP: Process only the first pole for debugging
        const polesToProcess = spidaPoles.slice(0, 1);
        polesToProcess.forEach((spidaPole, poleIndex) => {
            console.log(`--- Processing SPIDA Pole #${poleIndex + 1} ---`, spidaPole.label);
            console.log("SPIDA Pole Object:", JSON.parse(JSON.stringify(spidaPole)));

            operationNumber++;
            const firstDataRowForThisPole = reportData.length;
            let poleBlockRowCount = 0;

            // Define helper functions at the start of the pole processing scope

            // Improved wire descriptions based on usage group and client item
            function getWireDescription(wire) {
                if (!wire) return "Unknown";
                
                // Try to get description or size directly
                const clientItemDesc = safeAccess(wire, 'clientItem.description');
                if (clientItemDesc) return clientItemDesc;
                
                // Otherwise build description based on usage group and properties
                const usageGroup = safeAccess(wire, 'usageGroup', '').toUpperCase();
                const size = safeAccess(wire, 'clientItem.size', '');
                const owner = safeAccess(wire, 'owner.id', '');
                
                if (usageGroup.includes('PRIMARY')) return owner ? `${owner} Primary` : 'Primary';
                if (usageGroup === 'NEUTRAL') return 'Neutral';
                if (usageGroup === 'COMMUNICATION_BUNDLE' || usageGroup === 'COMMUNICATIONS') {
                    if (owner.toUpperCase().includes('CHARTER') || owner.toUpperCase().includes('SPECTRUM')) return 'Charter/Spectrum Fiber Optic';
                    if (owner.toUpperCase().includes('AT&T')) {
                        return (size && size.toLowerCase().includes('fiber')) || clientItemDesc && clientItemDesc.toLowerCase().includes('fiber') ? 'AT&T Fiber Optic Com' : 'AT&T Telco Com';
                    }
                    if (owner.toUpperCase().includes('CPS')) return 'CPS Supply Fiber'; // CPS Communications
                    return owner ? `${owner} Fiber Optic Com` : 'Fiber Optic Com'; // Generic Com
                }
                if (usageGroup === 'COMMUNICATION_SERVICE') {
                     if (owner.toUpperCase().includes('AT&T')) return 'AT&T Com Drop';
                     return owner ? `${owner} Com Drop` : 'Com Drop';
                }
                if (usageGroup === 'UTILITY_SERVICE' && owner.toUpperCase().includes('CPS')) return 'CPS Secondary Drop Loop';
                if (usageGroup.includes('STREET_LIGHT')) return owner ? `${owner} Street Light Drop` : 'Street Light Drop';
                if (usageGroup === 'RISER_EQUIPMENT' || safeAccess(wire, 'clientItem.type') === 'RISER') return owner ? `${owner} Riser` : 'Riser';
                if (usageGroup === 'ANCHOR_GUY_EQUIPMENT' || safeAccess(wire, 'clientItem.type') === 'GUY_ASSEMBLY') return owner ? `${owner} Guy` : 'Guy';

                if (owner && size) return `${owner} ${size}`;
                if (size) return size;
                if (owner) return `${owner} Cable`;
                return usageGroup || 'Unknown Attachment';
            }

            // Function to get mid-span data for a wire
            // TODO: Enhance this to use Katapult connection data for UG status
            function getMidSpanData(attachment, katapultConnectionForSpan) {
                // Check Katapult connection type first if available
                if (katapultConnectionForSpan) {
                    const connType = safeAccess(katapultConnectionForSpan, 'button', '').toLowerCase();
                    if (connType === 'underground_path') return "UG";
                }
                // Fallback to checking attachment description
                const description = attachment.description || '';
                if (description.toLowerCase().includes('underground') || description.toLowerCase().includes('ug')) {
                    return "UG";
                }
                return "NA";
            }
            
            // Improved span header formatting to match target
            function formatSpanHeader(wireEndPoint) {
                if (!wireEndPoint) return "Primary Span";

                if (wireEndPoint.type === "PREVIOUS_POLE") {
                    return "Backspan";
                }

                const directionDegrees = parseFloat(wireEndPoint.direction);
                let directionFormatted = '';

                if (isNaN(directionDegrees)) {
                    directionFormatted = wireEndPoint.direction || 'Unknown Direction'; 
                } else {
                    const directions = ["North", "North East", "East", "South East", "South", "South West", "West", "North West"];
                    const normalizedDegrees = (directionDegrees % 360 + 360) % 360;
                    const index = Math.round(normalizedDegrees / 45) % 8;
                    directionFormatted = directions[index];
                }
                
                const targetPoleLabel = wireEndPoint.structureLabel || 'Unknown Target';
                
                if (/^(?:PL|P\.|PO|\d+-PL)?\d+$/i.test(targetPoleLabel) || /^\d+$/.test(targetPoleLabel)) {
                     return `Ref (${directionFormatted}) to ${getCanonicalPoleID(targetPoleLabel, 'spida')}`;
                } else {
                    return `Ref (${directionFormatted}) to ${targetPoleLabel}`;
                }
            }

            const spidaPoleLabel = safeAccess(spidaPole, 'label', 'UNKNOWN_SPIDA_POLE');
            const canonicalPoleID = getCanonicalPoleID(spidaPoleLabel, 'spida');
            const katapultPoleData = katapultPoleMap.get(canonicalPoleID) || {};
            console.log("Canonical Pole ID:", canonicalPoleID);
            console.log("Matched Katapult Pole Data:", JSON.parse(JSON.stringify(katapultPoleData)));
            
            // --- POLE-LEVEL DATA ---
            const poleOwner = safeAccess(katapultPoleData, 'attributes.pole_owner.multi_added.0') || // Katapult preferred
                            safeAccess(katapultPoleData, 'attributes.pole_owner.button_added') ||
                            safeAccess(spidaPole, 'structure.pole.owner.id', 'NA');

            let poleStructure = 'NA';
            const spidaPoleHeightM = safeAccess(spidaPole, 'structure.pole.clientItem.height.value');
            if (spidaPoleHeightM) {
                poleStructure = convertToFeetInches(spidaPoleHeightM, safeAccess(spidaPole, 'structure.pole.clientItem.height.unit', 'METRE').toUpperCase());
                const poleClass = safeAccess(spidaPole, 'structure.pole.clientItem.classOfPole');
                const poleSpecies = safeAccess(spidaPole, 'structure.pole.clientItem.species');
                if (poleClass) poleStructure += `-${poleClass}`;
                if (poleSpecies) poleStructure += ` ${poleSpecies}`;
            }

            const recommendedDesign = (spidaPole.designs || []).find(d => d.name === "Recommended Design" || d.label === "Recommended Design") || 
                                   safeAccess(spidaPole, 'designs.1', {}); // Fallback to second design
            const measuredDesign = (spidaPole.designs || []).find(d => d.name === "Measured Design" || d.label === "Measured Design") || 
                                 safeAccess(spidaPole, 'designs.0', {}); // Fallback to first design
            console.log("Measured Design Object:", JSON.parse(JSON.stringify(measuredDesign)));
            console.log("Recommended Design Object:", JSON.parse(JSON.stringify(recommendedDesign)));

            const riserCount = (safeAccess(recommendedDesign, 'structure.equipments', []) || []).filter(eq => safeAccess(eq, 'clientItem.type') === 'RISER').length;
            const guyCount = (safeAccess(recommendedDesign, 'structure.guys', []) || []).length;
            
            let plaPercentage = "NA";
            let constructionGrade = "NA";

            // Revised logic for targetAnalysis
            let targetAnalysis = null;
            if (recommendedDesign && Array.isArray(recommendedDesign.analysis) && recommendedDesign.analysis.length > 0) {
                targetAnalysis = recommendedDesign.analysis.find(a =>
                    safeAccess(a, 'analysisCaseDetails.name', '').includes('Light - Grade C') || 
                    safeAccess(a, 'id', '').includes('Light - Grade C') // Check id as well from SPIDA example
                ) || recommendedDesign.analysis[0]; // Fallback to the first analysis in Recommended
            }
            // Fallback: if not found in recommended, try the spidaPole.analysis array directly
            if (!targetAnalysis && Array.isArray(spidaPole.analysis) && spidaPole.analysis.length > 0) {
                targetAnalysis = spidaPole.analysis.find(a =>
                    safeAccess(a, 'analysisCaseDetails.name', '').includes('Recommended') ||
                    safeAccess(a, 'analysisCaseDetails.name', '').includes('Light - Grade C')
                ) || spidaPole.analysis[spidaPole.analysis.length - 1];
            }

            console.log("Target SPIDA Analysis for PLA/Grade:", JSON.parse(JSON.stringify(targetAnalysis)));
            if (targetAnalysis) {
                const poleStressResult = (safeAccess(targetAnalysis, 'results', []) || []).find(r => r.component === "Pole" && r.analysisType === "STRESS");
                if (poleStressResult && typeof poleStressResult.actual === 'number') {
                    plaPercentage = formatPercentage(poleStressResult.actual);
                }
                constructionGrade = safeAccess(targetAnalysis, 'analysisCaseDetails.constructionGrade', 'NA');
            }
             // Fallback for PLA from Katapult (remains the same)
             if (plaPercentage === "NA") { 
                const katPLA = safeAccess(katapultPoleData, 'attributes.final_passing_capacity_%.assessment') ||
                               safeAccess(katapultPoleData, 'attributes.final_passing_capacity_%.auto_calced');
                if (katPLA) plaPercentage = formatPercentage(parseFloat(katPLA));
            }


            let heightLowestCom = "NA";
            let heightLowestCpsElec = "NA";
            // Logic for J & K: Use Katapult ATTACHMENT heights at the pole (not mid-span)
            // This matches the example image's apparent data for the first pole.
            if (katapultPoleData.attachments) {
                let minComH = Infinity, minCpsH = Infinity;
                Object.values(katapultPoleData.attachments).forEach(kAtt => {
                    const owner = safeAccess(kAtt, 'attributes.company_name.company_name', '').toUpperCase();
                    const hFt = parseFloat(safeAccess(kAtt, 'attributes.height_ft.assessment', 0));
                    const hIn = parseFloat(safeAccess(kAtt, 'attributes.height_in.assessment', 0));
                    const totalHFeet = hFt + (hIn / 12);
                    if (totalHFeet > 0) {
                        if (owner !== 'CPS ENERGY' && !owner.includes('POWER')) { // Basic check for COM
                            if (totalHFeet < minComH) minComH = totalHFeet;
                        } else if (owner === 'CPS ENERGY' || owner.includes('POWER')) { // Basic check for CPS Elec
                            if (totalHFeet < minCpsH) minCpsH = totalHFeet;
                        }
                    }
                });
                if (minComH !== Infinity) heightLowestCom = convertToFeetInches(minComH, 'FT');
                if (minCpsH !== Infinity) heightLowestCpsElec = convertToFeetInches(minCpsH, 'FT');
            }


            // --- ATTACHMENT CONSOLIDATION & ACTION ---
            const consolidatedAttachments = new Map();
            console.log("Initial (empty) consolidatedAttachments map:", consolidatedAttachments);

            function getAttachmentKey(owner, type) { // Simplified key for consolidation
                return `${String(owner).toUpperCase()}_${String(type).toUpperCase()}`;
            }
            
            // Process SPIDA Measured
            const spidaMeasuredAttachments = [...(safeAccess(measuredDesign, 'structure.wires', []) || []), ...(safeAccess(measuredDesign, 'structure.equipments', []) || [])];
            console.log("SPIDA Measured Attachments (raw from JSON):", JSON.parse(JSON.stringify(spidaMeasuredAttachments)));
            spidaMeasuredAttachments.forEach(att => {
                const owner = safeAccess(att, 'owner.id', 'Unknown');
                const type = getWireDescription(att); // Use full description for initial type
                const key = getAttachmentKey(owner, type);
                const heightM = safeAccess(att, 'attachmentHeight.value');
                
                if (!consolidatedAttachments.has(key)) {
                    consolidatedAttachments.set(key, {
                        description: type,
                        owner: owner,
                        spidaIDs: [],
                        state: 'measured_only', // Initial state
                        spidaMeasuredHeight: heightM ? convertToFeetInches(heightM, safeAccess(att, 'attachmentHeight.unit', 'METRE').toUpperCase()) : "NA",
                        spidaRecommendedHeight: "NA",
                        katapultExistingHeight: "NA",
                        proposedMidSpan: "NA" 
                    });
                }
                consolidatedAttachments.get(key).spidaIDs.push(att.id);
                if (heightM && consolidatedAttachments.get(key).spidaMeasuredHeight === "NA") { // Update if was NA
                     consolidatedAttachments.get(key).spidaMeasuredHeight = convertToFeetInches(heightM, safeAccess(att, 'attachmentHeight.unit', 'METRE').toUpperCase());
                }
            });

            // Process SPIDA Recommended
            const spidaRecommendedAttachments = [...(safeAccess(recommendedDesign, 'structure.wires', []) || []), ...(safeAccess(recommendedDesign, 'structure.equipments', []) || [])];
            console.log("SPIDA Recommended Attachments (raw from JSON):", JSON.parse(JSON.stringify(spidaRecommendedAttachments)));
            spidaRecommendedAttachments.forEach(att => {
                const owner = safeAccess(att, 'owner.id', 'Unknown');
                const type = getWireDescription(att);
                const key = getAttachmentKey(owner, type);
                const heightM = safeAccess(att, 'attachmentHeight.value');
                const recHeight = heightM ? convertToFeetInches(heightM, safeAccess(att, 'attachmentHeight.unit', 'METRE').toUpperCase()) : "NA";

                if (!consolidatedAttachments.has(key)) {
                    consolidatedAttachments.set(key, {
                        description: type,
                        owner: owner,
                        spidaIDs: [],
                        state: 'recommended_only', // New
                        spidaMeasuredHeight: "NA",
                        spidaRecommendedHeight: recHeight,
                        katapultExistingHeight: "NA",
                        proposedMidSpan: "NA"
                    });
                } else {
                    const existing = consolidatedAttachments.get(key);
                    existing.spidaRecommendedHeight = recHeight;
                    if (existing.state === 'measured_only') {
                        existing.state = (existing.spidaMeasuredHeight !== recHeight && recHeight !== "NA") ? 'modified' : 'existing';
                    }
                }
                consolidatedAttachments.get(key).spidaIDs.push(att.id);
                 if (heightM && consolidatedAttachments.get(key).spidaRecommendedHeight === "NA") {
                     consolidatedAttachments.get(key).spidaRecommendedHeight = recHeight;
                }
            });

            // Process Katapult Attachments (for existing height prioritization)
            let katapultAttachmentsRaw = safeAccess(katapultPoleData, 'attachments', null);
            if (!katapultAttachmentsRaw || Object.keys(katapultAttachmentsRaw).length === 0) {
                katapultAttachmentsRaw = safeAccess(katapultPoleData, 'attributes.attachments', {}); 
            }
            console.log("Katapult Attachments (raw from JSON for this pole):", JSON.parse(JSON.stringify(katapultAttachmentsRaw)));
            Object.values(katapultAttachmentsRaw).forEach(kAtt => {
                const owner = safeAccess(kAtt, 'attributes.company_name.company_name', 'Unknown');
                // Infer type from Katapult - this is tricky, needs robust mapping to SPIDA types
                // For now, let's try to match based on owner and a generic "cable" or description if possible
                // This part is complex and might need a more sophisticated Katapult type interpretation
                const kAttType = safeAccess(kAtt, 'attributes.attachment_type.button_added') || 
                                 safeAccess(kAtt, 'attributes.cable_type.button_added') ||
                                 'UnknownKatapultType'; // Needs better mapping
                
                // Attempt to find a matching key if possible, or rely on owner + a generic type.
                // This matching is simplified. A more robust approach would map Katapult types to SPIDA types.
                let keyToUpdate = null;
                consolidatedAttachments.forEach((value, key) => {
                    if (key.startsWith(owner.toUpperCase() + "_")) { // Simple owner match
                         // Further check if kAttType matches something in value.description
                        if (value.description.toLowerCase().includes(kAttType.toLowerCase()) || kAttType === 'UnknownKatapultType') {
                           keyToUpdate = key;
                        }
                    }
                });


                if (keyToUpdate) { // Only update if a plausible match is found
                    const heightFt = safeAccess(kAtt, 'attributes.height_ft.assessment');
                    const heightIn = safeAccess(kAtt, 'attributes.height_in.assessment');
                    if (heightFt && heightIn) {
                        const katHeight = convertToFeetInches(parseFloat(heightFt) + (parseFloat(heightIn)/12), 'FT');
                        consolidatedAttachments.get(keyToUpdate).katapultExistingHeight = katHeight;
                    }
                }
            });
            
            console.log("Final Consolidated Attachments Map (before generating rows):", JSON.parse(JSON.stringify(Array.from(consolidatedAttachments.entries()))));

            let attachmentAction = "( E )xisting";
            let hasInstall = false, hasRemoval = false, hasMod = false;
            consolidatedAttachments.forEach(att => {
                if (att.state === 'recommended_only') hasInstall = true;
                else if (att.state === 'measured_only') hasRemoval = true;
                else if (att.state === 'modified') hasMod = true;
            });

            if (hasInstall) attachmentAction = "( I )nstalling";
            else if (hasRemoval && !hasInstall) { // Only consider pure removal if no installs
                 // Check if anything remains or is modified
                let othersExistOrModified = false;
                consolidatedAttachments.forEach(att => {
                    if (att.state === 'existing' || att.state === 'modified') {
                        othersExistOrModified = true;
                    }
                });
                if (!othersExistOrModified && Array.from(consolidatedAttachments.values()).every(a => a.state === 'measured_only')) {
                     attachmentAction = "( R )emoving"; // All are to be removed
                } else if (hasRemoval && !othersExistOrModified && consolidatedAttachments.size === Array.from(consolidatedAttachments.values()).filter(a=>a.state === 'measured_only').length){
                     attachmentAction = "( R )emoving"; // Only removals left
                }
                 else {
                    attachmentAction = "( E )xisting"; // Some removed, others existing/modified
                }
            } else if (hasMod && !hasInstall && !hasRemoval) {
                attachmentAction = "( E )xisting"; // Modifications only
            }
            // Override by Katapult work_type
            const katWorkType = safeAccess(katapultPoleData, 'attributes.kat_work_type.button_added', '').toLowerCase();
            if (katWorkType === 'denied') attachmentAction = "( E )xisting (Denied)"; // Example for denied


            // --- ATTACHMENT ROWS ---
            const poleBlockRows = [];
            const spidaWireEndPoints = safeAccess(recommendedDesign, 'structure.wireEndPoints', []);

            if (spidaWireEndPoints.length === 0) { // No spans, list all recommended attachments under "Primary Span"
                const spanHeaderRow = new Array(15).fill(null);
                spanHeaderRow[11] = "Primary Span"; // L
                poleBlockRows.push(spanHeaderRow);
                poleBlockRowCount++;

                consolidatedAttachments.forEach(att => {
                    if (att.state === 'recommended_only' || att.state === 'existing' || att.state === 'modified') {
                        const row = new Array(15).fill(null);
                        row[11] = att.description; // L
                        row[12] = att.katapultExistingHeight !== "NA" ? att.katapultExistingHeight : att.spidaMeasuredHeight; // M
                        if (att.state === 'recommended_only') row[12] = ''; // No existing for new
                        row[13] = att.spidaRecommendedHeight; // N
                        row[14] = att.proposedMidSpan; // O - TODO: Improve getMidSpanData
                        poleBlockRows.push(row);
                        poleBlockRowCount++;
                    }
                });
                 if (poleBlockRows.length === 1 && poleBlockRowCount === 1) { // Only span header, no attachments
                    const noAttRow = new Array(15).fill(null);
                    noAttRow[11] = "No attachments";
                    poleBlockRows.push(noAttRow);
                    poleBlockRowCount++;
                }


            } else {
                spidaWireEndPoints.forEach(wep => {
                    const spanHeaderRow = new Array(15).fill(null);
                    spanHeaderRow[11] = formatSpanHeader(wep); // L
                    poleBlockRows.push(spanHeaderRow);
                    poleBlockRowCount++;
                    
                    const attachmentsOnThisSpan = new Set(); // To avoid duplicating same attachment on same span

                    (wep.wires || []).forEach(spidaWireIdOnSpan => {
                        consolidatedAttachments.forEach((att, key) => {
                            if (att.spidaIDs.includes(spidaWireIdOnSpan) && !attachmentsOnThisSpan.has(key)) {
                                if (att.state === 'recommended_only' || att.state === 'existing' || att.state === 'modified') {
                                    const row = new Array(15).fill(null);
                                    row[11] = att.description; // L
                                    row[12] = att.katapultExistingHeight !== "NA" ? att.katapultExistingHeight : att.spidaMeasuredHeight; // M
                                    if (att.state === 'recommended_only') row[12] = '';
                                    row[13] = att.spidaRecommendedHeight; // N
                                    // TODO: Pass wep or its derived Katapult connection to getMidSpanData
                                    // const katConnection = findKatapultConnectionForWep(wep, katapultPoleData, katapultData.connections);
                                    // row[14] = getMidSpanData(att, katConnection); // O
                                    row[14] = "NA"; // Placeholder for now
                                    poleBlockRows.push(row);
                                    poleBlockRowCount++;
                                    attachmentsOnThisSpan.add(key);
                                }
                            }
                        });
                    });
                    if (attachmentsOnThisSpan.size === 0 && poleBlockRows[poleBlockRows.length-1][11] === formatSpanHeader(wep) ) {
                        const noAttRow = new Array(15).fill(null);
                        noAttRow[11] = "No attachments on this span";
                         poleBlockRows.push(noAttRow);
                         poleBlockRowCount++;
                    }
                });
            }

            // Add placeholder if no attachment rows generated at all for pole
            if (poleBlockRowCount === 0) {
                const row = new Array(15).fill(null);
                row[11] = "NA"; row[12] = "NA"; row[13] = "NA"; row[14] = "NA";
                poleBlockRows.push(row);
                poleBlockRowCount++;
            }
            
            // --- FROM/TO POLE ---
            const fromPoleRow = new Array(15).fill(null);
            fromPoleRow[11] = "From Pole"; // L
            fromPoleRow[12] = canonicalPoleID; // M - As per spec
            poleBlockRows.push(fromPoleRow);
            poleBlockRowCount++;

            const toPoleRow = new Array(15).fill(null);
            toPoleRow[11] = "To Pole"; // L
            let toPoleID = "NA";
            if (katapultPoleData.katapultNodeId && katapultData.connections) {
                for (const conn of Object.values(katapultData.connections)) {
                    let targetNodeId = null;
                    if (conn.node_id_1 === katapultPoleData.katapultNodeId) targetNodeId = conn.node_id_2;
                    else if (conn.node_id_2 === katapultPoleData.katapultNodeId) targetNodeId = conn.node_id_1;
                    
                    if (targetNodeId) {
                        const targetKatPole = Object.values(katapultData.nodes || {}).find(n => n.katapultNodeId === targetNodeId || safeAccess(n, 'node_id') === targetNodeId);
                        if (!targetKatPole) { // If targetNodeId itself is the key in katapultData.nodes
                             const directNode = katapultData.nodes[targetNodeId];
                             if(directNode) {
                                const poleNum = safeAccess(directNode, 'attributes.PoleNumber.assessment') || safeAccess(directNode, 'attributes.pole_tag.tagtext');
                                if (poleNum) toPoleID = getCanonicalPoleID(typeof poleNum === 'string' ? poleNum : safeAccess(poleNum, 'tagtext', String(poleNum)), 'katapult');
                             }
                        } else if (targetKatPole) {
                            const poleNum = safeAccess(targetKatPole, 'attributes.PoleNumber.assessment') || safeAccess(targetKatPole, 'attributes.pole_tag.tagtext');
                            if (poleNum) toPoleID = getCanonicalPoleID(typeof poleNum === 'string' ? poleNum : safeAccess(poleNum, 'tagtext', String(poleNum)), 'katapult');
                        }
                        if (toPoleID !== "NA") break; 
                    }
                }
            }
            if (toPoleID === "NA") { // Fallback to SPIDA
                const firstOutboundWep = spidaWireEndPoints.find(wep => wep.type !== "PREVIOUS_POLE" && wep.structureLabel);
                if (firstOutboundWep) toPoleID = getCanonicalPoleID(firstOutboundWep.structureLabel, 'spida');
            }
            toPoleRow[12] = toPoleID; // M
            poleBlockRows.push(toPoleRow);
            poleBlockRowCount++;

            // --- ADD POLE BLOCK TO REPORT & MERGE ---
            // Add pole-level data to the first row of the block
            if (poleBlockRows.length > 0) {
                const firstRowInBlock = poleBlockRows[0];
                firstRowInBlock[0] = operationNumber;
                firstRowInBlock[1] = attachmentAction;
                firstRowInBlock[2] = poleOwner;
                firstRowInBlock[3] = canonicalPoleID;
                firstRowInBlock[4] = poleStructure;
                firstRowInBlock[5] = formatYesNoCount(riserCount);
                firstRowInBlock[6] = formatYesNoCount(guyCount);
                firstRowInBlock[7] = plaPercentage;
                firstRowInBlock[8] = constructionGrade;
                firstRowInBlock[9] = heightLowestCom;
                firstRowInBlock[10] = heightLowestCpsElec;
            }
            
            reportData.push(...poleBlockRows);

            if (poleBlockRowCount > 1) { // Merge if more than one row for this pole (including From/To)
                 // End merge at the row just BEFORE "From Pole"
                const endMergeRow = firstDataRowForThisPole + poleBlockRowCount - 3; // -2 for From/To, -1 for 0-indexed offset
                if (endMergeRow >= firstDataRowForThisPole) {
                    for (let col = 0; col < 11; col++) { // A-K
                        merges.push({
                            s: { r: firstDataRowForThisPole, c: col },
                            e: { r: endMergeRow, c: col }
                        });
                    }
                }
            }
        });

        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(reportData);
        ws['!merges'] = merges;

        // Set column widths
        ws['!cols'] = [
            { wch: 10 },  // A: Operation Number
            { wch: 20 },  // B: Attachment Action
            { wch: 15 },  // C: Pole Owner
            { wch: 20 },  // D: Pole #
            { wch: 25 },  // E: Pole Structure
            { wch: 18 },  // F: Proposed Riser
            { wch: 18 },  // G: Proposed Guy
            { wch: 18 },  // H: PLA
            { wch: 18 },  // I: Construction Grade
            { wch: 20 },  // J: Height Lowest Com
            { wch: 20 },  // K: Height Lowest CPS Elec
            { wch: 35 },  // L: Attacher Desc
            { wch: 15 },  // M: Existing Height
            { wch: 15 },  // N: Proposed Height
            { wch: 20 }   // O: Mid-Span Proposed
        ];

        // Enable text wrapping for specific cells
        const cellB1 = XLSX.utils.encode_cell({r: 0, c: 1});
        const cellO2 = XLSX.utils.encode_cell({r: 1, c: 14});
        [cellB1, cellO2].forEach(cellRef => {
            if (ws[cellRef]) {
                ws[cellRef].s = {
                    alignment: { wrapText: true, vertical: 'center', horizontal: 'center' }
                };
            }
        });

        // Create workbook and add worksheet
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Make Ready Report");

        // Generate and download Excel file
        XLSX.writeFile(wb, "Make_Ready_Report.xlsx");
    }
});
