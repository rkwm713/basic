document.addEventListener('DOMContentLoaded', () => {
    const jsonFileInput = document.getElementById('jsonFileInput');
    const generateReportBtn = document.getElementById('generateReportBtn');
    const reportTableContainer = document.getElementById('reportTableContainer');

    generateReportBtn.addEventListener('click', () => {
        if (jsonFileInput.files.length === 0) {
            alert('Please select a Katapult JSON file first.');
            return;
        }
        const file = jsonFileInput.files[0];
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const katapultData = JSON.parse(event.target.result);
                generateReport(katapultData);
            } catch (error) {
                console.error('Error parsing JSON file:', error);
                reportTableContainer.innerHTML = `<p style="color: red;">Error parsing JSON file: ${error.message}</p>`;
            }
        };

        reader.onerror = () => {
            console.error('Error reading file.');
            reportTableContainer.innerHTML = `<p style="color: red;">Error reading file.</p>`;
        };

        reader.readAsText(file);
    });

    function generateReport(data) {
        if (!data.nodes || !data.connections) {
            reportTableContainer.innerHTML = `<p style="color: red;">Invalid Katapult JSON format: Missing 'nodes' or 'connections'.</p>`;
            return;
        }

        const table = document.createElement('table');
        table.innerHTML = `
            <thead>
                <tr>
                    <th rowspan="3">Operation Number</th>
                    <th rowspan="3" class="text-wrap">Attachment Action:\\n( I )nstalling\\n( R )emoving\\n( E )xisting</th>
                    <th rowspan="3">Pole Owner</th>
                    <th rowspan="3">Pole #</th>
                    <th rowspan="3">Pole Structure</th>
                    <th rowspan="3">Proposed Riser (Yes/No) &</th>
                    <th rowspan="3">Proposed Guy (Yes/No) &</th>
                    <th rowspan="3">PLA (%) with proposed attachment</th>
                    <th rowspan="3">Construction Grade of Analysis</th>
                    <th colspan="2" class="merged-header">Existing Mid-Span Data</th>
                    <th colspan="4" class="merged-header">Make Ready Data</th>
                </tr>
                <tr>
                    <th rowspan="2">Height Lowest Com</th>
                    <th rowspan="2">Height Lowest CPS Electrical</th>
                    <th colspan="3" class="merged-header">Attachment Height</th>
                    <th rowspan="2" class="text-wrap">Mid-Span\\n(same span as existing)</th>
                </tr>
                <tr>
                    <th>Attacher Description</th>
                    <th>Existing</th>
                    <th>Proposed</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        `;
        const tbody = table.querySelector('tbody');
        // let operationNumber = 1; // Removed, will be sourced from scid

        for (const nodeId in data.nodes) {
            if (Object.hasOwnProperty.call(data.nodes, nodeId)) {
                const node = data.nodes[nodeId];
                const attributes = node.attributes || {};

                // PRIORITY 2: Implement Node Type Filtering
                const nodeType = getAttributeValue(attributes.node_type, ['-Imported', 'button_added', 'one']); // Added 'one' as a common key
                const allowedNodeTypes = ["pole", "Power", "Power Transformer", "Joint", "Joint Transformer"];
                if (!nodeType || !allowedNodeTypes.includes(nodeType)) {
                    continue; // Skip this node if its type is not in the allowed list
                }

                const allPhotosData = data.photos || {};
                const allTracesData = data.traces && data.traces.trace_data ? data.traces.trace_data : {};

                const operationNumberValue = getAttributeValue(attributes.scid, ['-Imported']) || 
                                           getAttributeValue(attributes.OP_number, ['-Imported']) || 
                                           'NA'; // Fallback to OP_number then NA

                // --- Extract Pole-Level Data (Columns A-K) ---
                const poleOwnerCompany = getAttributeValue(attributes.pole_owner, ['multi_added', 'button_added']) || ''; // Used by midspan funcs
                const poleOwner = poleOwnerCompany || 'NA';
                const poleNumber = getAttributeValue(attributes.PoleNumber, ['assessment', '-Imported']) || getAttributeValue(attributes.pole_tag, ['tagtext']) || 'NA';
                const poleHeight = getAttributeValue(attributes.height, ['one']) || '';
                const poleClass = getAttributeValue(attributes.pole_class, ['one']) || '';
                const poleSpecies = getPoleSpecies(attributes.birthmark_brand) || '';
                const poleStructure = `${poleHeight}${poleHeight && (poleClass || poleSpecies) ? '-' : ''}${poleClass}${poleClass && poleSpecies ? '-' : ''}${poleSpecies}` || 'NA';

                const proposedRiser = getProposedEquipment(node, data.connections, 'riser'); // Simplified
                const proposedGuy = getProposedEquipment(node, data.connections, 'guy'); // Simplified

                const pla = formatPercentage(getAttributeValue(attributes.final_passing_capacity_p, ['assessment', '-ONzZigRJczUNfA6wSoG', '-ONzZihf2oHcPH25VAt5'])) || // Check multiple possible keys
                              formatPercentage(getAttributeValue(attributes['final_passing_capacity_%'], ['assessment', '-ONzZigRJczUNfA6wSoG', '-ONzZihf2oHcPH25VAt5'])) || 'NA';


                const constructionGrade = getAttributeValue(attributes.pole_class, ['one']) || 'NA'; // Per mapping doc
                const existingMidSpanCom = getMidSpanHeight(nodeId, data.connections, allPhotosData, allTracesData, 'com', poleOwnerCompany) || '';
                const existingMidSpanCps = getMidSpanHeight(nodeId, data.connections, allPhotosData, allTracesData, 'cps_electrical', poleOwnerCompany) || '';

                // Attachment Action - Simplified for now
                let attachmentAction = '( E )xisting';
                if (attributes.mr_state && getAttributeValue(attributes.mr_state, ['auto_calced']) === 'Needs MR Review') {
                    attachmentAction = '( I )nstalling / ( R )emoving'; // Placeholder
                }


                // --- Extract Attachment-Level Data (Columns L-O) ---
                const attachments = [];
                const processAttachmentItem = (item) => {
                    const { existingHeightInches, mrMoveInches, traceId } = getPoleAttachmentDetails(item);

                    let desc = 'NA';
                    if (traceId && allTracesData[traceId]) {
                        const traceInfo = allTracesData[traceId];
                        const company = traceInfo.company || '';
                        const cableType = traceInfo.cable_type || '';
                        
                        if (company && cableType) {
                            desc = `${company} ${cableType}`;
                        } else if (company) {
                            desc = company;
                        } else if (cableType) {
                            desc = cableType;
                        }
                    }
                    // Fallback for description if traceId lookup fails or is incomplete
                    // This might need adjustment based on whether 'wire' and 'equipment' objects have alternative attributes
                    if (desc === 'NA' || desc.trim() === '') {
                        // Example: if item has 'item_company_name' or 'item_type'
                        // const companyNameAttr = item.item_company_name; 
                        // const attachmentTypeAttr = item.item_type;
                        // if (companyNameAttr) desc = companyNameAttr;
                        // else if (attachmentTypeAttr) desc = attachmentTypeAttr;
                        // For now, if trace lookup fails, it remains 'NA' or what was found
                    }
                    if (desc.trim() === '') desc = 'NA';

                    const makeReadyExistingHeight = formatHeightFtIn(existingHeightInches);
                    let makeReadyProposedHeight = '';
                    if (mrMoveInches != null && !isNaN(mrMoveInches) && existingHeightInches != null) {
                        makeReadyProposedHeight = formatHeightFtIn(existingHeightInches + mrMoveInches);
                    } else if (existingHeightInches != null) {
                        makeReadyProposedHeight = makeReadyExistingHeight;
                    }

                    const midSpanProposedHeight = getMidspanProposedHeight(traceId, mrMoveInches, nodeId, data.connections, allPhotosData, allTracesData);

                    attachments.push({
                        desc,
                        existingHeight: makeReadyExistingHeight,
                        proposedHeight: makeReadyProposedHeight,
                        midspanProposed: midSpanProposedHeight
                    });
                };

                if (node.photofirst_data) {
                    if (node.photofirst_data.wire) {
                        for (const wireId in node.photofirst_data.wire) {
                            if (Object.hasOwnProperty.call(node.photofirst_data.wire, wireId)) {
                                processAttachmentItem(node.photofirst_data.wire[wireId]);
                            }
                        }
                    }
                    if (node.photofirst_data.equipment) {
                        for (const equipId in node.photofirst_data.equipment) {
                            if (Object.hasOwnProperty.call(node.photofirst_data.equipment, equipId)) {
                                processAttachmentItem(node.photofirst_data.equipment[equipId]);
                            }
                        }
                    }
                }

                if (attachments.length === 0) {
                    attachments.push({ desc: 'NA', existingHeight: '', proposedHeight: '', midspanProposed: '' });
                }

                // --- Create Rows ---
                const numRowsForPole = attachments.length + 2;

                attachments.forEach((att, index) => {
                    const row = tbody.insertRow();
                    if (index === 0) {
                        addCell(row, operationNumberValue, numRowsForPole);
                        addCell(row, attachmentAction, numRowsForPole);
                        addCell(row, poleOwner, numRowsForPole);
                        addCell(row, poleNumber, numRowsForPole);
                        addCell(row, poleStructure, numRowsForPole);
                        addCell(row, proposedRiser, numRowsForPole);
                        addCell(row, proposedGuy, numRowsForPole);
                        addCell(row, pla, numRowsForPole);
                        addCell(row, constructionGrade, numRowsForPole);
                        addCell(row, existingMidSpanCom, numRowsForPole);
                        addCell(row, existingMidSpanCps, numRowsForPole);
                    }
                    addCell(row, att.desc);
                    addCell(row, att.existingHeight);
                    addCell(row, att.proposedHeight);
                    addCell(row, att.midspanProposed);
                });

                // Add "From Pole" / "To Pole" rows
                const fromPoleRow = tbody.insertRow();
                if (attachments.length === 0) { // This condition might need re-evaluation if the default attachment is always added
                     addCell(fromPoleRow, operationNumberValue, 2, true);
                     addCell(fromPoleRow, attachmentAction, 2, true);
                     addCell(fromPoleRow, poleOwner, 2, true);
                     addCell(fromPoleRow, poleNumber, 2, true);
                     addCell(fromPoleRow, poleStructure, 2, true);
                     addCell(fromPoleRow, proposedRiser, 2, true);
                     addCell(fromPoleRow, proposedGuy, 2, true);
                     addCell(fromPoleRow, pla, 2, true);
                     addCell(fromPoleRow, constructionGrade, 2, true);
                     addCell(fromPoleRow, existingMidSpanCom, 2, true);
                     addCell(fromPoleRow, existingMidSpanCps, 2, true);
                } else {
                    // Cells A-K are already spanned by the attachment logic
                }
                addCell(fromPoleRow, 'From Pole');
                addCell(fromPoleRow, poleNumber); 
                addCell(fromPoleRow, '');
                addCell(fromPoleRow, '');

                const toPoleRow = tbody.insertRow();
                const connectedPoleInfo = getConnectedPole(nodeId, data.connections, data.nodes);
                addCell(toPoleRow, 'To Pole');
                addCell(toPoleRow, connectedPoleInfo.poleNumber || 'NA');
                addCell(toPoleRow, '');
                addCell(toPoleRow, '');

            }
        }

        reportTableContainer.innerHTML = ''; 
        reportTableContainer.appendChild(table);
    }

    function inchesToFeet(inchesValue) {
        if (inchesValue == null || inchesValue === '') return "";
        const numInches = parseFloat(inchesValue);
        if (isNaN(numInches)) return "";
        return (numInches / 12).toFixed(1); // Returns decimal feet string
    }

    function formatHeightFtIn(totalInches) {
        if (totalInches == null || totalInches === '' || isNaN(parseFloat(totalInches))) {
            return '';
        }
        const inchesNum = parseFloat(totalInches);
        const feet = Math.floor(inchesNum / 12);
        const inches = Math.round(inchesNum % 12); 
        return `${feet}'-${inches}"`;
    }

    // Helper to get details for a pole-mounted attachment (wire or equipment item)
    function getPoleAttachmentDetails(item) { // item is a wire or equipment object
        let existingHeightInches = null;
        if (item && item._measured_height != null) {
            const height = parseFloat(item._measured_height);
            if (!isNaN(height)) {
                existingHeightInches = height; // Assuming _measured_height is already in inches
            }
        }

        let mrMoveInches = null;
        if (item && item.mr_move != null) {
            const move = parseFloat(item.mr_move);
            if (!isNaN(move)) {
                mrMoveInches = move; // Assuming mr_move is already in inches
            }
        }
        
        const traceId = (item && item._trace) ? item._trace : null;

        return { existingHeightInches, mrMoveInches, traceId };
    }

    function addCell(row, text, rowspan = 1, isPlaceholderForPoleBlock = false) {
        const cell = row.insertCell();
        cell.textContent = text;
        if (rowspan > 1) {
            cell.rowSpan = rowspan;
            if (!isPlaceholderForPoleBlock && row.parentElement.rows.length > 1 && row.rowIndex > 0) {
                 // Avoid setting rowspan again
            }
        }
    }

    function getAttributeValue(attributeObj, keys) {
        if (!attributeObj) return null;
        for (const key of keys) {
            if (attributeObj[key] !== undefined && attributeObj[key] !== null) {
                if (typeof attributeObj[key] === 'object' && attributeObj[key] !== null) {
                    const subKeys = Object.keys(attributeObj[key]);
                    if (subKeys.length > 0) return attributeObj[key][subKeys[0]];
                }
                return attributeObj[key];
            }
        }
        for (const key of keys) {
            if (key.includes('.')) {
                const parts = key.split('.');
                let current = attributeObj;
                for (const part of parts) {
                    if (current && current[part] !== undefined) {
                        current = current[part];
                    } else {
                        current = null;
                        break;
                    }
                }
                if (current !== null) return current;
            }
        }
        return null;
    }

    function getPoleSpecies(birthmarkBrand) {
        if (!birthmarkBrand) return '';
        for (const id in birthmarkBrand) {
            if (birthmarkBrand[id] && birthmarkBrand[id]['pole_species*']) {
                const speciesCode = birthmarkBrand[id]['pole_species*'];
                if (speciesCode === 'SPC') return 'Southern Pine';
                return speciesCode; 
            }
        }
        return '';
    }

    function formatHeight(ft, inches) { // This function seems unused now, formatHeightFtIn is used
        if (ft === null && inches === null) return '';
        const ftVal = ft || 0;
        const inVal = inches || 0;
        return `${ftVal}'-${inVal}"`;
    }

    function formatPercentage(value) {
        if (value === null || value === undefined || value === '') return '';
        const num = parseFloat(value);
        if (isNaN(num)) return '';
        return num.toFixed(2) + '%';
    }

    function getProposedEquipment(node, connections, type) {
        let count = 0;
        if (node.attributes) {
            for (const key in node.attributes) {
                if (key.toLowerCase().includes(type) && getAttributeValue(node.attributes[key], ['button_added', 'assessment', 'one'])) {
                    count++; 
                }
            }
        }
        if (type === 'guy') {
            for (const connId in connections) {
                const conn = connections[connId];
                if (conn.node_id_1 === node.id || conn.node_id_2 === node.id) { // node.id might be an issue, should be nodeId
                    if (getAttributeValue(conn.attributes?.connection_type, ['button_added'])?.toLowerCase().includes('guy') ||
                        conn.button?.toLowerCase().includes('anchor')) {
                        count++;
                    }
                }
            }
        }
        return count > 0 ? `YES (${count})` : 'NO';
    }

    function getMidSpanHeight(currentNodeId, connections, allPhotosData, allTracesData, type, poleOwnerCompany) {
        let minHeightInches = Infinity;
        if (!connections || !allPhotosData || !allTracesData) return "";

        for (const connId in connections) {
            if (Object.hasOwnProperty.call(connections, connId)) {
                const conn = connections[connId];
                if ((conn.node_id_1 === currentNodeId || conn.node_id_2 === currentNodeId) &&
                    getAttributeValue(conn.attributes?.connection_type, ['button_added']) === 'aerial cable') {
                    if (conn.sections) {
                        for (const sectionId in conn.sections) {
                            if (Object.hasOwnProperty.call(conn.sections, sectionId)) {
                                const section = conn.sections[sectionId];
                                if (section.photos) {
                                    for (const photoKey in section.photos) {
                                        const photoId = (typeof section.photos[photoKey] === 'string') ? section.photos[photoKey] : photoKey;
                                        const photoData = allPhotosData[photoId];

                                        if (photoData && photoData.photofirst_data && photoData.photofirst_data.wire) {
                                            for (const wireId in photoData.photofirst_data.wire) {
                                                const wire = photoData.photofirst_data.wire[wireId];
                                                if (wire._trace && wire._measured_height != null) {
                                                    const traceInfo = allTracesData[wire._trace];
                                                    if (traceInfo) {
                                                        let matchesType = false;
                                                        if (type === 'com') {
                                                            if (traceInfo.company && poleOwnerCompany && traceInfo.company !== poleOwnerCompany) {
                                                                matchesType = true;
                                                            }
                                                        } else if (type === 'cps_electrical') {
                                                            if (traceInfo.company && poleOwnerCompany && traceInfo.company === poleOwnerCompany &&
                                                                (traceInfo.cable_type === 'Primary' || traceInfo.cable_type === 'Neutral' || traceInfo.cable_type === 'Secondary')) { 
                                                                matchesType = true;
                                                            }
                                                        }

                                                        if (matchesType) {
                                                            const currentHeightInches = parseFloat(wire._measured_height);
                                                            if (!isNaN(currentHeightInches)) {
                                                                minHeightInches = Math.min(minHeightInches, currentHeightInches);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        return minHeightInches === Infinity ? "" : inchesToFeet(minHeightInches);
    }

    function getMidspanProposedHeight(poleAttachmentTraceId, poleAttachmentMrMoveInches, currentNodeId, connections, allPhotosData, allTracesData) {
        if (poleAttachmentTraceId == null || poleAttachmentMrMoveInches == null || isNaN(parseFloat(poleAttachmentMrMoveInches))) {
            return "";
        }
        if (!connections || !allPhotosData || !allTracesData) return "";

        const mrMoveInchesNum = parseFloat(poleAttachmentMrMoveInches);

        for (const connId in connections) {
            if (Object.hasOwnProperty.call(connections, connId)) {
                const conn = connections[connId];
                if ((conn.node_id_1 === currentNodeId || conn.node_id_2 === currentNodeId) &&
                    getAttributeValue(conn.attributes?.connection_type, ['button_added']) === 'aerial cable') {
                    if (conn.sections) {
                        for (const sectionId in conn.sections) {
                            if (Object.hasOwnProperty.call(conn.sections, sectionId)) {
                                const section = conn.sections[sectionId];
                                if (section.photos) {
                                    for (const photoKey in section.photos) {
                                        const photoId = (typeof section.photos[photoKey] === 'string') ? section.photos[photoKey] : photoKey;
                                        const photoData = allPhotosData[photoId];
                                        if (photoData && photoData.photofirst_data && photoData.photofirst_data.wire) {
                                            for (const wireId in photoData.photofirst_data.wire) {
                                                const wire = photoData.photofirst_data.wire[wireId];
                                                if (wire._trace === poleAttachmentTraceId && wire._measured_height != null) {
                                                    const midspanMeasuredHeightInches = parseFloat(wire._measured_height);
                                                    if (!isNaN(midspanMeasuredHeightInches)) {
                                                        const proposedMidspanHeightInches = midspanMeasuredHeightInches + mrMoveInchesNum;
                                                        return inchesToFeet(proposedMidspanHeightInches);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        return ""; 
    }

    function getConnectedPole(currentNodeId, connectionsData, nodesData) {
        for (const connId in connectionsData) {
            const conn = connectionsData[connId];
            let targetNodeId = null;
            if (conn.node_id_1 === currentNodeId) {
                targetNodeId = conn.node_id_2;
            } else if (conn.node_id_2 === currentNodeId) {
                targetNodeId = conn.node_id_1;
            }

            if (targetNodeId && nodesData[targetNodeId]) {
                const targetNode = nodesData[targetNodeId];
                const poleNum = getAttributeValue(targetNode.attributes?.PoleNumber, ['assessment', '-Imported']) ||
                                getAttributeValue(targetNode.attributes?.pole_tag, ['tagtext']);
                return { poleNumber: poleNum, connectionType: getAttributeValue(conn.attributes?.connection_type, ['button_added']) || conn.button };
            }
        }
        return {};
    }
});
