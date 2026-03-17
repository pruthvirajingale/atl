const express = require("express");
const mongoose = require("mongoose");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const ExcelJS = require("exceljs");
const cors = require("cors");
const path = require("path");

const app = express();
const SECRET = process.env.JWT_SECRET || "change_this_secret";

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

mongoose.connect("mongodb+srv://ingalepruthviraj50_db_user:iIwYmLipv2IxChqH@cluster99.apwbb2y.mongodb.net/")
  .then(() => console.log("MongoDB connected"))
  .catch(console.error);

/* ── Models ── */

const User = mongoose.model("User", new mongoose.Schema({
  name:     { type: String, required: true },
  email:    { type: String, required: true, unique: true, lowercase: true },
  password: { type: String, required: true },
  role:     { type: String, enum: ["admin", "hod", "class_teacher", "subject_teacher"], required: true },
  departmentId:     { type: mongoose.Schema.Types.ObjectId, ref: "Department", default: null },
  assignedSubjects: [{ subjectId: mongoose.Schema.Types.ObjectId, divisionId: mongoose.Schema.Types.ObjectId }]
}));

const Department = mongoose.model("Department", new mongoose.Schema({
  name:  { type: String, required: true },
  code:  { type: String, required: true, unique: true, uppercase: true },
  hodId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null },
  // years[].divisions[].rollRange
  years: [{
    year: { type: Number, enum: [1, 2, 3] },
    divisions: [{
      name: String,
      rollRange: { start: Number, end: Number }
    }]
  }]
}));

const Subject = mongoose.model("Subject", new mongoose.Schema({
  name:         { type: String, required: true },
  code:         { type: String, required: true },
  departmentId: { type: mongoose.Schema.Types.ObjectId, ref: "Department", required: true },
  year:         { type: Number, enum: [1, 2, 3], required: true },
  semester:     { type: Number, enum: [1, 2], required: true },
  type:         { type: String, enum: ["TH", "PR", "TH+PR"], required: true },
  divisionId:   { type: mongoose.Schema.Types.ObjectId, required: true },
  divisionName: { type: String, required: true },
  teacherId:    { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null }
}));

const Attendance = mongoose.model("Attendance", new mongoose.Schema({
  studentRoll:  { type: String, required: true },
  studentName:  { type: String, required: true },
  subjectId:    { type: mongoose.Schema.Types.ObjectId, ref: "Subject", required: true },
  departmentId: { type: mongoose.Schema.Types.ObjectId, required: true },
  divisionId:   { type: mongoose.Schema.Types.ObjectId, required: true },
  year:         Number,
  semester:     Number,
  date:         { type: String, required: true },
  markedBy:     { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null }
}));

Attendance.schema.index({ studentRoll: 1, subjectId: 1, date: 1 }, { unique: true });

/* ── Middleware ── */

function auth(...roles) {
  return (req, res, next) => {
    try {
      const token = req.headers.authorization?.split(" ")[1];
      if (!token) return res.status(401).json({ error: "No token." });
      req.user = jwt.verify(token, SECRET);
      if (roles.length && !roles.includes(req.user.role))
        return res.status(403).json({ error: "Access denied." });
      next();
    } catch {
      res.status(401).json({ error: "Invalid token." });
    }
  };
}

/* ── Helpers ── */

const hashPw  = pw => bcrypt.hash(pw, 10);
const checkPw = (pw, hash) => bcrypt.compare(pw, hash);

function makeToken(user) {
  return jwt.sign(
    { id: user._id, role: user.role, departmentId: user.departmentId, assignedSubjects: user.assignedSubjects || [] },
    SECRET, { expiresIn: "7d" }
  );
}

function getDateRange(range) {
  const to = new Date(), from = new Date();
  if (range === "week")   from.setDate(to.getDate() - 6);
  if (range === "2weeks") from.setDate(to.getDate() - 13);
  if (range === "month")  from.setMonth(to.getMonth() - 1);
  return { from: from.toISOString().split("T")[0], to: to.toISOString().split("T")[0] };
}

function fmtDate(iso) {
  const [y, m, d] = iso.split("-");
  return `${d} ${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][+m-1]} ${y.slice(2)}`;
}

/* ── Auth ── */

app.get("/auth/check-admin", async (req, res) => {
  const exists = await User.exists({ role: "admin" });
  res.json({ exists: !!exists });
});

app.post("/auth/signup", async (req, res) => {
  try {
    const { name, email, password } = req.body;
    if (!name || !email || !password) return res.status(400).json({ error: "All fields required." });
    if (await User.findOne({ email })) return res.status(409).json({ error: "Email already in use." });
    const user = await User.create({ name, email, password: await hashPw(password), role: "admin" });
    res.json({ token: makeToken(user), name: user.name, role: user.role });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/auth/login", async (req, res) => {
  try {
    const { email, password } = req.body;
    const user = await User.findOne({ email });
    if (!user || !(await checkPw(password, user.password)))
      return res.status(401).json({ error: "Invalid email or password." });
    res.json({ token: makeToken(user), name: user.name, role: user.role, departmentId: user.departmentId });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ── Users (admin) ── */

app.get("/users", auth("admin", "hod"), async (req, res) => {
  const filter = {};
  if (req.query.role)         filter.role = req.query.role;
  // HOD can only see teachers in their own department
  if (req.user.role === "hod") filter.departmentId = req.user.departmentId;
  else if (req.query.departmentId) filter.departmentId = req.query.departmentId;
  const users = await User.find(filter, "-password").lean();
  res.json(users);
});

app.post("/users", auth("admin", "hod"), async (req, res) => {
  try {
    const { name, email, password, role } = req.body;
    if (!name || !email || !password || !role) return res.status(400).json({ error: "All fields required." });
    if (req.user.role === "hod" && !["class_teacher", "subject_teacher"].includes(role))
      return res.status(403).json({ error: "HOD can only create teacher accounts." });
    if (await User.findOne({ email })) return res.status(409).json({ error: "Email in use." });
    const departmentId = req.user.role === "hod" ? req.user.departmentId : req.body.departmentId;
    const user = await User.create({ name, email, password: await hashPw(password), role, departmentId });
    res.json({ _id: user._id, name: user.name, role: user.role });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.patch("/users/:id", auth("admin"), async (req, res) => {
  const user = await User.findByIdAndUpdate(req.params.id, req.body, { new: true, select: "-password" });
  res.json(user);
});

app.delete("/users/:id", auth("admin", "hod"), async (req, res) => {
  const user = await User.findById(req.params.id);
  if (!user) return res.status(404).json({ error: "User not found." });
  // HOD can only delete teachers in their own department
  if (req.user.role === "hod" && user.departmentId?.toString() !== req.user.departmentId?.toString())
    return res.status(403).json({ error: "Access denied." });
  await User.findByIdAndDelete(req.params.id);
  res.json({ message: "Deleted." });
});

/* ── Departments (admin) ── */

app.get("/departments", auth(), async (req, res) => {
  const depts = await Department.find().populate("hodId", "name email").lean();
  res.json(depts);
});

app.get("/departments/:id", auth(), async (req, res) => {
  const dept = await Department.findById(req.params.id).populate("hodId", "name email").lean();
  res.json(dept);
});

app.post("/departments", auth("admin"), async (req, res) => {
  try {
    const dept = await Department.create(req.body);
    if (req.body.hodId) await User.findByIdAndUpdate(req.body.hodId, { departmentId: dept._id });
    res.json(dept);
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.patch("/departments/:id", auth("admin"), async (req, res) => {
  const dept = await Department.findByIdAndUpdate(req.params.id, req.body, { new: true });
  if (req.body.hodId) await User.findByIdAndUpdate(req.body.hodId, { departmentId: dept._id });
  res.json(dept);
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
        filter.subject = { $regex: `^${subject}_(TH|PR)$` };
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
    const { subjectId, divisionId, year, semester, departmentId, range, date, from, to } = req.query;
    const f = {};

    if (req.user.role === "subject_teacher")
      f.subjectId = { $in: (req.user.assignedSubjects || []).map(a => a.subjectId) };
    else if (req.user.role === "class_teacher")
      f.departmentId = req.user.departmentId;
    else if (req.user.role === "hod")
      f.departmentId = req.user.departmentId;

    if (subjectId)    f.subjectId    = subjectId;
    if (divisionId)   f.divisionId   = divisionId;
    if (departmentId && req.user.role === "admin") f.departmentId = departmentId;
    if (year)         f.year         = +year;
    if (semester)     f.semester     = +semester;

    if (date) f.date = date;
    else if (from || to) f.date = { ...(from && { $gte: from }), ...(to && { $lte: to }) };
    else if (range) { const r = getDateRange(range); f.date = { $gte: r.from, $lte: r.to }; }

    res.json(await Attendance.find(f).sort({ date: 1 }).lean());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Toggle single record
app.patch("/attendance", auth("admin", "hod", "subject_teacher"), async (req, res) => {
  try {
    const { studentRoll, subjectId, date, present, studentName } = req.body;
    if (present) {
      const s = await Subject.findById(subjectId).lean();
      await Attendance.updateOne(
        { studentRoll, subjectId, date },
        { $set: { studentName, divisionId: s.divisionId, departmentId: s.departmentId,
                  year: s.year, semester: s.semester, markedBy: req.user.id } },
        { upsert: true }
      );
    } else {
      await Attendance.deleteOne({ studentRoll, subjectId, date });
    }
    res.json({ message: "Updated." });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ── Export Excel ── */

app.get("/export/:subjectId", auth(), async (req, res) => {
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
      records = await Attendance.find({ subject: { $regex: `^${subject}_(TH|PR)$` } }).lean();
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
      subject: { $regex: `^${base}_(TH|PR)$` },
      date: { $gte: from, $lte: to }
    }).lean();

    if (!records.length) return res.status(404).send("No data.");

    const dates = [...new Set(records.map(r => r.date))].sort();
    const students = {};
    records.forEach(r => {
      students[r.studentRoll] ??= { roll: r.studentRoll, name: r.studentName, att: {} };
      students[r.studentRoll].att[r.date] = true;
    });

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(`${subject.code}-${subject.divisionName}`);
    ws.addRow(["Roll", "Name", ...dates.map(fmtDate), "Present", "Total", "%"]);
    ws.getRow(1).font = { bold: true };

    Object.values(students)
      .sort((a, b) => +a.roll - +b.roll)
      .forEach(s => {
        const attended = dates.filter(d => s.att[d]).length;
        ws.addRow([s.roll, s.name, ...dates.map(d => s.att[d] ? "P" : "A"),
          attended, dates.length, dates.length ? ((attended / dates.length) * 100).toFixed(1) + "%" : "N/A"]);
      });

    res.setHeader("Content-Disposition",
      `attachment; filename=${subject.code}_${subject.divisionName}_Y${subject.year}S${subject.semester}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) { res.status(500).send("Export failed."); }
});

/* ── Start ── */

app.get("/health", (_, res) => res.json({ ok: true }));
app.get("/", (req, res) => res.sendFile(path.join(__dirname, "public", "admin.html")));
app.use((_, res) => res.status(404).send("Not found."));
app.listen(process.env.PORT || 3000, () => console.log("Server running."));