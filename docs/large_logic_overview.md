Detailed Logic Mapping for Make-Ready Report Generation
Report Structure Understanding
The Excel report has a pole-centric structure where:
1.	Primary rows contain pole-level information (columns A-K)
2.	Sub-rows contain attachment-specific information for each attacher on that pole (columns L-O)
3.	From/To Pole rows appear between poles showing span information
Column-by-Column Logic Mapping
Column A: Operation Number
•	Source: Manual/External - not from JSON files
•	Logic: Sequential numbering (1, 2, 3, etc.) for each pole
•	Implementation: Generate incrementally during report creation
Column B: Attachment Action (I/R/E)
•	Primary Sources: SPIDAcalc (design comparison), Katapult (proposed traces)
•	Logic: Determine action for the primary attacher (Charter/Spectrum)
SPIDAcalc Logic:
// Compare Measured vs Recommended designs
"designs": [
  {
    "label": "Measured Design",
    "layerType": "Measured",
    "structure": {
      "wires": [/* existing wires */],
      "equipments": [/* existing equipment */]
    }
  },
  {
    "label": "Recommended Design", 
    "layerType": "Recommended",
    "structure": {
      "wires": [/* proposed wires */],
      "equipments": [/* proposed equipment */]
    }
  }
]
Katapult Logic:
// Check for proposed traces
"traces": {
  "trace_data": {
    "trace_id": {
      "company": "Charter",
      "proposed": true  // Indicates installing
    }
  }
}

// Check for make-ready moves
"nodes": {
  "node_id": {
    "photos": {
      "photo_id": {
        "photofirst_data": {
          "wire": {
            "wire_id": {
              "mr_move": 24  // Non-zero indicates relocation
            }
          }
        }
      }
    }
  }
}
Decision Logic:
•	(I)nstalling: Charter wire exists in Recommended but not Measured OR Katapult trace has proposed: true
•	(R)emoving: Charter wire exists in Measured but not Recommended
•	(E)xisting: Charter wire exists in both with same properties
Column C: Pole Owner
•	Primary Source: SPIDAcalc Measured Design
•	Fallback: Katapult attributes
SPIDAcalc Path:
"designs[?(@.layerType=='Measured')].structure.pole.owner.id"
// Example: "CPS Energy"
Katapult Path:
"nodes[node_id].attributes.pole_owner.multi_added"
// Example: "CPS Energy"
Column D: Pole
•	Primary Source: SPIDAcalc location label
•	Mapping Required: Correlate SPIDAcalc poles with Katapult nodes
SPIDAcalc Path:
"leads[*].locations[*].label"
// Example: "1-PL410620" -> extract "PL410620"
Katapult Correlation Paths:
"nodes[node_id].attributes.PoleNumber.-Imported"
"nodes[node_id].attributes.DLOC_number.-Imported"
"nodes[node_id].attributes.electric_pole_tag.assessment"
Column E: Pole Structure
•	Primary Source: SPIDAcalc client data lookup
•	Fallback: Katapult attributes
SPIDAcalc Logic:
// 1. Get pole reference from structure
"designs[?(@.layerType=='Measured')].structure.pole.clientItem.id"

// 2. Lookup in client data
"clientData.poles[?(@.aliases[*].id == pole_ref)].species"
"clientData.poles[?(@.aliases[*].id == pole_ref)].classOfPole"

// Result: "Southern Pine 4" -> "40-4 Southern Pine"
Katapult Path:
"nodes[node_id].attributes.pole_species.one"
"nodes[node_id].attributes.pole_class.one"
 
Column F: Proposed Riser (Yes/No)
•	Primary Source: SPIDAcalc Recommended Design
•	Check For: Equipment with type "RISER"
SPIDAcalc Logic:
"designs[?(@.layerType=='Recommended')].structure.equipments"
// Check if any equipment has:
"clientItem.type.name": "RISER"
Katapult Path:
"nodes[node_id].attributes.riser.button_added"
// "Yes" or "No"
Column G: Proposed Guy (Yes/No)
•	Primary Source: SPIDAcalc Recommended Design
•	Check For: Guys array or proposed guy traces
SPIDAcalc Logic:
"designs[?(@.layerType=='Recommended')].structure.guys"
// Non-empty array indicates proposed guy
Katapult Logic:
// Check for proposed guys
"nodes[node_id].attributes.guying[guy_id].proposed": true

// Or check notes for guy installation
"nodes[node_id].attributes.kat_MR_notes[note_id]"
// Text containing "install down guy"
Column H: PLA (%) with Proposed Attachment
•	Source: SPIDAcalc Recommended Design analysis results
Path:
"designs[?(@.layerType=='Recommended')].structure.pole.stressRatio"
// Multiply by 100 for percentage

// OR from analysis results:
"designs[?(@.layerType=='Recommended')].analysis[*].results"
// Find result where component is pole and unit is "PERCENT"
Column I: Construction Grade of Analysis
•	Source: SPIDAcalc analysis case details
Path:
"designs[?(@.layerType=='Recommended')].analysis[*].analysisCaseDetails.constructionGrade"
// Example: "C"
Columns J & K: Existing Mid-Span Data
Column J: Height Lowest Com
•	Primary Source: SPIDAcalc Measured Design span wires
•	Logic: Find minimum mid-span height of communication wires
SPIDAcalc Logic:
// 1. Identify communication wires in Measured Design
"designs[?(@.layerType=='Measured')].structure.wires"
// Where wire usage groups include "COMMUNICATION"

// 2. Get mid-span heights
"midspanHeight.value" // Convert from metres to feet
Katapult Logic:
// Check mid-span sections in connections
"connections[connection_id].sections[section_id].annotations"
// Find communication wires and their measured heights
Column K: Height Lowest CPS Electrical
•	Similar logic to Column J but for electrical wires
•	Filter: Owner = "CPS Energy" AND usage groups include "PRIMARY", "SECONDARY", "NEUTRAL"
 
Columns L-O: Make Ready Data (Per Attacher)
These columns repeat for each attacher on the pole. The report shows multiple rows per pole:
Column L: Attacher Description
•	Logic: Describe each attachment (wire/equipment) on the pole
Examples from Excel:
•	"Neutral"
•	"CPS Supply Fiber"
•	"Charter/Spectrum Fiber Optic"
•	"AT&T Fiber Optic Com"
SPIDAcalc Logic:
// For wires
"structure.wires[*].owner.id" + " " + "clientItem.description"

// For equipment  
"structure.equipments[*].owner.id" + " " + "clientItem.type.name"
Column M: Attachment Height - Existing
•	Source: Current height of attachment in Measured Design
SPIDAcalc Path:
"designs[?(@.layerType=='Measured')].structure.wires[*].attachmentHeight.value"
// Convert from metres to feet and format as "XX'-YY""
Katapult Path:
"nodes[node_id].attributes.equipment[equipment_id].attachment_height_ft"
// Already in feet-inches format
Column N: Attachment Height - Proposed
•	Source: New height in Recommended Design OR calculated from mr_move
SPIDAcalc Path:
"designs[?(@.layerType=='Recommended')].structure.wires[*].attachmentHeight.value"
Katapult Calculation:
// If mr_move exists:
existing_height + mr_move_value = proposed_height
Column O: Mid-Span - Proposed
•	Source: SPIDAcalc Recommended Design mid-span calculations
Path:
"designs[?(@.layerType=='Recommended')].structure.wires[*].midspanHeight.value"
// Convert from metres to feet
 
Special Rows: From/To Pole Information
Between pole sections, the report includes span information:
Logic:
// SPIDAcalc: Find connected poles through wireEndPoints
"structure.wireEndPoints[*].externalId"

// Katapult: Find span connections  
"connections[connection_id].node_id_1" 
"connections[connection_id].node_id_2"
Report Generation Algorithm
1.	Initialize: Set up Excel template with headers
2.	Correlate Poles: Map SPIDAcalc locations to Katapult nodes
3.	For Each Pole: 
o	Extract pole-level data (Columns A-K)
o	Find all attachments on pole (both existing and proposed)
o	For each attachment, extract attachment data (Columns L-O)
o	Add span information between poles
4.	Format: Apply proper formatting (feet-inches, percentages, etc.)
Key Conversion Functions Needed
1.	Height Conversion: Metres to Feet-Inches format
2.	Pole Number Extraction: Remove prefixes from SPIDAcalc labels
3.	Wire Type Mapping: Map technical descriptions to readable names
4.	Owner Normalization: Standardize owner names between systems

