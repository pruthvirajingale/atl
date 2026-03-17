const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");

const app = express();

app.use(cors()); // allows all origins
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

mongoose.connect(
"mongodb+srv://ingalepruthviraj50_db_user:iIwYmLipv2IxChqH@cluster99.apwbb2y.mongodb.net/"
)
.then(() => console.log("MongoDB Atlas connected"))
.catch(err => console.error(err));

const attendanceSchema = new mongoose.Schema({
  roll: { type: String, required: true },
  name: { type: String, required: true },
  subject: { type: String, required: true }, // NIS_TH, NIS_PR
  date: {
    type: String,
    required: true,
    match: /^\d{4}-\d{2}-\d{2}$/
  }
});

attendanceSchema.index(
  { roll: 1, subject: 1, date: 1 },
  { unique: true }
);

const Attendance = mongoose.model("Attendance", attendanceSchema);

 function formatExcelDate(isoDate) {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${d} ${months[Number(m) - 1]} ${y.slice(2)}`;
}

function getBaseSubject(subject) {
  return subject.replace(/_(TH|PR)$/, "");
}

function getType(subject) {
  if (subject.endsWith("_TH")) return "TH";
  if (subject.endsWith("_PR")) return "PR";
  return "";
}

/*Routes*/
app.get("/health", (req, res) => {
  res.json({ status: "OK" });
});

/*Mark Attendance*/
app.post("/attendance", async (req, res) => {
  try {
    const { roll, name, subject } = req.body;
    if (!roll || !name || !subject) {
      return res.status(400).json({ error: "Roll, name, subject required" });
    }

    const today = new Date().toISOString().split("T")[0];

    await Attendance.create({
      roll: roll.trim(),
      name: name.trim(),
      subject: subject.trim(),
      date: today
    });

    res.json({ message: "Attendance saved" });
  } catch (err) {
    if (err.code === 11000) {
      return res.status(409).json({ error: "Attendance already marked" });
    }
    console.error(err);
    res.status(500).json({ error: "Failed to save attendance" });
  }
});

app.get("/admin/attendance", async (req, res) => {
  try {
    const { subject, range } = req.query;
    let filter = {};

    // Filter by subject
    if (subject) {
      if (subject.endsWith("_TH") || subject.endsWith("_PR")) {
        filter.subject = subject;
      } else {
        filter.subject = { $in: [`${subject}_TH`, `${subject}_PR`] };
      }
    }

    // Calculate date range relative to today
    const today = new Date();
    let start = new Date();

    if (range === "week") start.setDate(today.getDate() - 6);      // last 7 days
    else if (range === "2weeks") start.setDate(today.getDate() - 13); // last 14 days
    else if (range === "month") start.setMonth(today.getMonth() - 1); // last month

    const from = start.toISOString().split("T")[0];
    const to = today.toISOString().split("T")[0];

    filter.date = { $gte: from, $lte: to }; // filter between from and to

    const data = await Attendance.find(filter).sort({ date: 1 }).lean();
    res.json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to load attendance" });
  }
});

app.patch("/admin/attendance", async (req, res) => {
  try {
    const { roll, subject, date, status, name } = req.body;

    if (!roll || !subject || !date || !status) {
      return res.status(400).json({ error: "Missing fields" });
    }

    if (status === "Present") {
      await Attendance.updateOne(
        { roll, subject, date },
        { $set: { name } },
        { upsert: true }
      );
    } else {
      await Attendance.deleteOne({ roll, subject, date });
    }

    res.json({ message: "Attendance updated" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Update failed" });
  }
});

/* ======================= Export Excel with Division & % ======================= */
app.get("/export-excel/:subject", async (req, res) => {
  try {
    const subject = req.params.subject;
    const division = req.query.division || "A"; // default Co-A if not specified

    // Fetch all records for this subject
    let records;
    if (subject.endsWith("_TH") || subject.endsWith("_PR")) {
      // Single TH or PR
      records = await Attendance.find({ subject }).lean();
    } else {
      // Combined TH+PR
      records = await Attendance.find({ subject: { $in: [`${subject}_TH`, `${subject}_PR`] } }).lean();
    }

    if (!records.length) return res.status(404).send("No data");

    // Filter by division roll ranges
    if (division === "A") records = records.filter(r => +r.roll >= 1 && +r.roll <= 63);
    if (division === "B") records = records.filter(r => +r.roll >= 64 && +r.roll <= 126);

    // Create unique lecture dates (or date+type for combined)
    const lectureKeys = [...new Set(
      records.map(r => r.subject.endsWith("_TH") || r.subject.endsWith("_PR")
        ? `${r.date}_${r.subject.endsWith("_TH") ? "TH" : "PR"}`
        : r.date
      )
    )].sort();

    // Map students
    const students = {};
    records.forEach(r => {
      students[r.roll] ??= { roll: r.roll, name: r.name, attendance: {} };
      const key = r.subject.endsWith("_TH") || r.subject.endsWith("_PR")
        ? `${r.date}_${r.subject.endsWith("_TH") ? "TH" : "PR"}`
        : r.date;
      students[r.roll].attendance[key] = true;
    });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(`${subject} ${division}`);

    // Header row
    ws.addRow(["Roll No", "Name", ...lectureKeys.map(k => {
      const [d, t] = k.split("_");
      return t ? `${formatExcelDate(d)} (${t})` : formatExcelDate(d);
    }), "Attendance %"]);
    ws.getRow(1).font = { bold: true };

    // Add student rows
    Object.values(students).forEach(s => {
      const attendedCount = lectureKeys.filter(k => s.attendance[k]).length;
      const total = lectureKeys.length;
      const percent = total ? ((attendedCount / total) * 100).toFixed(1) + "%" : "No Lectures";

      ws.addRow([
        s.roll,
        s.name,
        ...lectureKeys.map(k => s.attendance[k] ? "Present" : "Absent"),
        percent
      ]);
    });

    res.setHeader("Content-Disposition", `attachment; filename=${subject}_${division}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Export failed");
  }
});


app.get("/export-combined/:subject", async (req, res) => {
  try {
    const base = req.params.subject;
    const range = req.query.range;

    const today = new Date();
    let start = new Date();
    if (range === "week") start.setDate(today.getDate() - 6);
    else if (range === "2weeks") start.setDate(today.getDate() - 13);
    else if (range === "month") start.setMonth(today.getMonth() - 1);

    const from = start.toISOString().split("T")[0];
    const to = today.toISOString().split("T")[0];

    // Get records for this subject + date range
    const records = await Attendance.find({
      subject: { $in: [`${base}_TH`, `${base}_PR`] },
      date: { $gte: from, $lte: to }
    }).lean();

    if (!records.length) return res.status(404).send("No data");

    // Collect all lecture keys (date + TH/PR)
    const lectureKeys = [...new Set(records.map(r => `${r.date}_${r.subject.endsWith("_TH") ? "TH" : "PR"}`))].sort();

    // Map students
    const students = {};
    records.forEach(r => {
      students[r.roll] ??= { roll: r.roll, name: r.name, attendance: {} };
      students[r.roll].attendance[`${r.date}_${r.subject.endsWith("_TH") ? "TH" : "PR"}`] = "Present";
    });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(`${base} TH+PR`);

    // Header
    ws.addRow([
      "Roll No",
      "Name",
      ...lectureKeys.map(k => {
        const [d, t] = k.split("_");
        const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        const [y, m, day] = d.split("-");
        return `${day} ${months[+m-1]} ${y.slice(2)} (${t})`;
      }),
      "Attendance %"
    ]);
    ws.getRow(1).font = { bold: true };

    // Add student rows
    Object.values(students).forEach(s => {
      const rowData = lectureKeys.map(k => s.attendance[k] || "Absent");
      const attended = rowData.filter(v => v === "Present").length;
      const percent = lectureKeys.length === 0 ? 0 : ((attended / lectureKeys.length) * 100).toFixed(1);
      ws.addRow([s.roll, s.name, ...rowData, percent + "%"]);
    });

    // Send Excel
    res.setHeader("Content-Disposition", `attachment; filename=${base}_combined.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Combined export failed");
  }
});

const PORT = process.env.PORT || 3000;
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "admin.html"));
});
app.use((req, res) => res.status(404).send("File not found"));
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));