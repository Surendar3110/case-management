from flask import Flask, request, render_template_string, jsonify
import pandas as pd
import os

app = Flask(__name__)

DATA_FILE = "uploaded_data.xlsx"

# ---------------------------------------------------------
# HTML FRONT END
# ---------------------------------------------------------

HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>ZF Case Management</title>

    <style>

        * {
            box-sizing: border-box;
        }

        body {
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            background: #f4f6f9;
            color: #181818;
        }

        .header {
            background: #032d60;
            color: white;
            padding: 16px 30px;
            font-size: 20px;
            font-weight: bold;
        }

        .container {
            max-width: 900px;
            margin: 35px auto;
        }

        .card {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,.15);
            padding: 28px;
        }

        .title {
            font-size: 22px;
            font-weight: bold;
            margin-bottom: 25px;
        }

        .field {
            margin-bottom: 24px;
        }

        label {
            display: block;
            font-size: 13px;
            margin-bottom: 7px;
            color: #444;
            font-weight: 600;
        }

        select {
            width: 100%;
            height: 42px;
            border: 1px solid #c9c9c9;
            border-radius: 4px;
            padding: 0 12px;
            font-size: 14px;
            background: white;
        }

        select:focus {
            outline: none;
            border-color: #0176d3;
            box-shadow: 0 0 3px #0176d3;
        }

        .button {
            background: #0176d3;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 10px 20px;
            cursor: pointer;
            font-size: 14px;
        }

        .button:hover {
            background: #014486;
        }

        .upload-area {
            background: #f8f9fb;
            border: 1px dashed #aaa;
            padding: 18px;
            margin-bottom: 30px;
            border-radius: 5px;
        }

        .status {
            margin-top: 12px;
            font-size: 13px;
            color: #2e844a;
        }

        .result {
            margin-top: 25px;
            background: #f3f3f3;
            padding: 15px;
            border-radius: 5px;
        }

    </style>
</head>

<body>

<div class="header">
    ZF Case Management
</div>

<div class="container">

    <div class="card">

        <div class="title">
            Case Information
        </div>

        <!-- UPLOAD -->

        <div class="upload-area">

            <form id="uploadForm" enctype="multipart/form-data">

                <label>
                    Upload Case Master Data
                </label>

                <input
                    type="file"
                    id="file"
                    name="file"
                    accept=".xlsx,.xls,.csv"
                    required
                >

                <button
                    type="submit"
                    class="button"
                    style="margin-top:12px;"
                >
                    Upload Data
                </button>

            </form>

            <div id="uploadStatus" class="status"></div>

        </div>


        <!-- CASE REASON -->

        <div class="field">

            <label>
                Case Reason
            </label>

            <select id="caseReason">
                <option value="">--None--</option>
            </select>

        </div>


        <!-- VEHICLE TYPE -->

        <div class="field">

            <label>
                Vehicle Type
            </label>

            <select id="vehicleType">
                <option value="">--None--</option>
            </select>

        </div>


        <!-- SUB CASE -->

        <div class="field">

            <label>
                Sub Case Reason
            </label>

            <select id="subCase">
                <option value="">--None--</option>
            </select>

        </div>


        <div class="result">

            <b>Selected Values</b>

            <p id="result">
                No selection
            </p>

        </div>


        <!-- MANUAL ENTRY -->

        <div class="title" style="margin-top:20px; font-size:18px;">
            Add / Enter Values Manually
        </div>

        <div class="field">

            <label>
                Case Reason (manual)
            </label>

            <input id="manualCase" type="text" style="width:100%; height:36px; padding:6px 10px;" />

        </div>

        <div class="field">

            <label>
                Vehicle Type (manual)
            </label>

            <input id="manualVehicle" type="text" style="width:100%; height:36px; padding:6px 10px;" />

        </div>

        <div class="field">

            <label>
                Sub Case Reason (manual)
            </label>

            <input id="manualSub" type="text" style="width:100%; height:36px; padding:6px 10px;" />

        </div>

        <div style="margin-top:8px;">
            <button id="addRecord" class="button">Add Record</button>
            <span id="addStatus" class="status" style="margin-left:12px"></span>
        </div>

    </div>

</div>


<script>

let data = [];


/* ----------------------------------------------------
   LOAD DATA
---------------------------------------------------- */

async function loadData() {

    const response = await fetch("/data");

    data = await response.json();

    populateCaseReasons();

}


/* ----------------------------------------------------
   CASE REASON
---------------------------------------------------- */

function populateCaseReasons() {

    const select = document.getElementById("caseReason");

    select.innerHTML =
        '<option value="">--None--</option>';

    const values = [...new Set(
        data.map(x => x.case_reason)
    )].filter(Boolean).sort();

    values.forEach(value => {

        const option =
            document.createElement("option");

        option.value = value;
        option.textContent = value;

        select.appendChild(option);

    });

}


/* ----------------------------------------------------
   VEHICLE TYPE
---------------------------------------------------- */

function populateVehicleTypes() {

    const caseReason =
        document.getElementById("caseReason").value;

    const select =
        document.getElementById("vehicleType");

    select.innerHTML =
        '<option value="">--None--</option>';

    document.getElementById("subCase").innerHTML =
        '<option value="">--None--</option>';

    if (!caseReason)
        return;

    const values = [...new Set(

        data
            .filter(x =>
                x.case_reason === caseReason
            )
            .map(x =>
                x.vehicle_type
            )

    )].filter(Boolean).sort();


    values.forEach(value => {

        const option =
            document.createElement("option");

        option.value = value;
        option.textContent = value;

        select.appendChild(option);

    });

}


/* ----------------------------------------------------
   SUB CASE
---------------------------------------------------- */

function populateSubCases() {

    const caseReason =
        document.getElementById("caseReason").value;

    const vehicle =
        document.getElementById("vehicleType").value;

    const select =
        document.getElementById("subCase");

    select.innerHTML =
        '<option value="">--None--</option>';

    if (!caseReason || !vehicle)
        return;


    const values = [...new Set(

        data
            .filter(x =>
                x.case_reason === caseReason &&
                x.vehicle_type === vehicle
            )
            .map(x =>
                x.sub_case
            )

    )].filter(Boolean).sort();


    values.forEach(value => {

        const option =
            document.createElement("option");

        option.value = value;
        option.textContent = value;

        select.appendChild(option);

    });

}


/* ----------------------------------------------------
   DISPLAY RESULT
---------------------------------------------------- */

function showResult() {

    const caseReason =
        document.getElementById("caseReason").value;

    const vehicle =
        document.getElementById("vehicleType").value;

    const subCase =
        document.getElementById("subCase").value;


    document.getElementById("result").innerHTML =

        "<b>Case Reason:</b> " +
        (caseReason || "--None--") +

        "<br><b>Vehicle Type:</b> " +
        (vehicle || "--None--") +

        "<br><b>Sub Case Reason:</b> " +
        (subCase || "--None--");

}


/* ----------------------------------------------------
   EVENTS
---------------------------------------------------- */

document
    .getElementById("caseReason")
    .addEventListener("change", function() {

        populateVehicleTypes();
        showResult();

    });


document
    .getElementById("vehicleType")
    .addEventListener("change", function() {

        populateSubCases();
        showResult();

    });


document
    .getElementById("subCase")
    .addEventListener("change", showResult);


/* ----------------------------------------------------
   UPLOAD
---------------------------------------------------- */

document
    .getElementById("uploadForm")
    .addEventListener("submit", async function(e) {

        e.preventDefault();

        const formData =
            new FormData(this);

        const response =
            await fetch("/upload", {

                method: "POST",
                body: formData

            });

        const result =
            await response.json();

        document.getElementById(
            "uploadStatus"
        ).textContent = result.message;

        if (result.success) {

            await loadData();

            document.getElementById(
                "caseReason"
            ).value = "";

            document.getElementById(
                "vehicleType"
            ).innerHTML =
                '<option value="">--None--</option>';

            document.getElementById(
                "subCase"
            ).innerHTML =
                '<option value="">--None--</option>';

        }

    });


/* ----------------------------------------------------
   ADD RECORD (MANUAL ENTRY)
---------------------------------------------------- */

document
    .getElementById("addRecord")
    .addEventListener("click", async function(e) {

        e.preventDefault();

        const caseReason =
            document.getElementById("manualCase").value.trim();

        const vehicle =
            document.getElementById("manualVehicle").value.trim();

        const subCase =
            document.getElementById("manualSub").value.trim();

        if (!caseReason && !vehicle && !subCase) {

            document.getElementById("addStatus").textContent =
                "Please enter at least one value.";

            return;

        }

        const payload = {
            case_reason: caseReason,
            vehicle_type: vehicle,
            sub_case: subCase
        };

        document.getElementById("addStatus").textContent = "Saving...";

        try {

            const res = await fetch("/add_record", {

                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)

            });

            const result = await res.json();

            document.getElementById("addStatus").textContent = result.message;

            if (result.success) {

                // reset inputs
                document.getElementById("manualCase").value = "";
                document.getElementById("manualVehicle").value = "";
                document.getElementById("manualSub").value = "";

                await loadData();

                populateCaseReasons();

            }

        } catch (err) {

            document.getElementById("addStatus").textContent = "Error saving record.";

        }

    });

/* Initial load */

loadData();

</script>

</body>
</html>
"""


# ---------------------------------------------------------
# LOAD DATA
# ---------------------------------------------------------

def read_data():

    if not os.path.exists(DATA_FILE):

        return []

    try:

        if DATA_FILE.endswith(".csv"):

            df = pd.read_csv(DATA_FILE)

        else:

            df = pd.read_excel(DATA_FILE)

        df.columns = [
            str(c).strip().lower()
            for c in df.columns
        ]

        required = [
            "case reason",
            "vehicle type",
            "sub case reason"
        ]

        if all(x in df.columns for x in required):

            df = df.rename(columns={
                "case reason": "case_reason",
                "vehicle type": "vehicle_type",
                "sub case reason": "sub_case"
            })

        else:

            return []

        df = df.fillna("")

        return df[
            [
                "case_reason",
                "vehicle_type",
                "sub_case"
            ]
        ].to_dict(orient="records")

    except Exception as e:

        print("Error:", e)

        return []


# ---------------------------------------------------------
# HOME
# ---------------------------------------------------------

@app.route("/")
def home():

    return render_template_string(HTML)


# ---------------------------------------------------------
# RETURN DATA
# ---------------------------------------------------------

@app.route("/data")
def get_data():

    return jsonify(read_data())


# ---------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------

@app.route("/upload", methods=["POST"])
def upload():

    file = request.files.get("file")

    if not file:

        return jsonify({
            "success": False,
            "message": "Please select a file."
        })

    filename = file.filename.lower()

    if not (
        filename.endswith(".xlsx")
        or filename.endswith(".xls")
        or filename.endswith(".csv")
    ):

        return jsonify({
            "success": False,
            "message": "Only Excel or CSV files are allowed."
        })


    file.save(DATA_FILE)

    records = read_data()

    if not records:

        return jsonify({
            "success": False,
            "message":
            "File uploaded, but required columns were not found."
        })


    return jsonify({

        "success": True,

        "message":
        f"Data uploaded successfully. {len(records)} records loaded."

    })



@app.route("/add_record", methods=["POST"])
def add_record():

    try:

        payload = request.get_json() or {}

        case_reason = (payload.get("case_reason") or "").strip()
        vehicle_type = (payload.get("vehicle_type") or "").strip()
        sub_case = (payload.get("sub_case") or "").strip()

        # Build a dataframe row
        row = {
            "case_reason": case_reason,
            "vehicle_type": vehicle_type,
            "sub_case": sub_case
        }

        # If file exists, read it; otherwise create new df
        if os.path.exists(DATA_FILE):

            try:

                if DATA_FILE.endswith('.csv'):

                    df = pd.read_csv(DATA_FILE)

                else:

                    df = pd.read_excel(DATA_FILE)

                # Normalize column names if needed
                df.columns = [str(c).strip().lower() for c in df.columns]

                # Rename if original labels present
                df = df.rename(columns={
                    'case reason': 'case_reason',
                    'vehicle type': 'vehicle_type',
                    'sub case reason': 'sub_case'
                })

            except Exception:

                df = pd.DataFrame(columns=["case_reason", "vehicle_type", "sub_case"])

        else:

            df = pd.DataFrame(columns=["case_reason", "vehicle_type", "sub_case"])

        # Append row
        df = df.append(row, ignore_index=True)

        # Save back to Excel (overwrite)
        try:

            df.to_excel(DATA_FILE, index=False)

        except Exception:

            # fallback to CSV
            df.to_csv(DATA_FILE.replace('.xlsx', '.csv'), index=False)

        return jsonify({

            "success": True,

            "message": "Record added successfully."

        })

    except Exception as e:

        print("Add record error:", e)

        return jsonify({

            "success": False,

            "message": "Failed to add record."

        })


# ---------------------------------------------------------
# RUN
# ---------------------------------------------------------

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )