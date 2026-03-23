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

/* ── Database Connection ── */

let isConnected = false;

async function connectDB() {
  if (isConnected) return;
  if (mongoose.connection.readyState >= 1) {
    isConnected = true;
    return;
  }
  await mongoose.connect(process.env.MONGODB_URI || "mongodb+srv://ingalepruthviraj50_db_user:iHKXlcgpv9DLG1xS@cluster99.apwbb2y.mongodb.net/", {
    serverSelectionTimeoutMS: 5000,
    bufferCommands: false,
  });
  isConnected = true;
}

// Connect to DB on every request (required for Vercel serverless)
app.use(async (req, res, next) => {
  try {
    await connectDB();
    next();
  } catch (err) {
    console.error("DB connection error:", err);
    res.status(500).json({ error: "Database connection failed." });
  }
});

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

const attendanceSchema = new mongoose.Schema({
  studentRoll:  { type: String, required: true },
  studentName:  { type: String, required: true },
  subjectId:    { type: mongoose.Schema.Types.ObjectId, ref: "Subject", required: true },
  departmentId: { type: mongoose.Schema.Types.ObjectId, required: true },
  divisionId:   { type: mongoose.Schema.Types.ObjectId, required: true },
  year:         Number,
  semester:     Number,
  date:         { type: String, required: true },
  lectureType:  { type: String, enum: ["TH", "PR"], default: "TH" },
  markedBy:     { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null }
});

attendanceSchema.index({ studentRoll: 1, subjectId: 1, date: 1, lectureType: 1 }, { unique: true });

const Attendance = mongoose.model("Attendance", attendanceSchema);

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
  try {
    const exists = await User.exists({ role: "admin" });
    res.json({ exists: !!exists });
  } catch (e) { res.status(500).json({ error: e.message }); }
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
  try {
    const filter = {};
    if (req.query.role)         filter.role = req.query.role;
    if (req.user.role === "hod") filter.departmentId = req.user.departmentId;
    else if (req.query.departmentId) filter.departmentId = req.query.departmentId;
    const users = await User.find(filter, "-password").lean();
    res.json(users);
  } catch (e) { res.status(500).json({ error: e.message }); }
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
  try {
    const user = await User.findByIdAndUpdate(req.params.id, req.body, { new: true, select: "-password" });
    res.json(user);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete("/users/:id", auth("admin", "hod"), async (req, res) => {
  try {
    const user = await User.findById(req.params.id);
    if (!user) return res.status(404).json({ error: "User not found." });
    if (req.user.role === "hod" && user.departmentId?.toString() !== req.user.departmentId?.toString())
      return res.status(403).json({ error: "Access denied." });
    await User.findByIdAndDelete(req.params.id);
    res.json({ message: "Deleted." });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ── Departments (admin) ── */

app.get("/departments", auth(), async (req, res) => {
  try {
    const depts = await Department.find().populate("hodId", "name email").lean();
    res.json(depts);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get("/departments/:id", auth(), async (req, res) => {
  try {
    const dept = await Department.findById(req.params.id).populate("hodId", "name email").lean();
    res.json(dept);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/departments", auth("admin"), async (req, res) => {
  try {
    const dept = await Department.create(req.body);
    if (req.body.hodId) await User.findByIdAndUpdate(req.body.hodId, { departmentId: dept._id });
    res.json(dept);
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.patch("/departments/:id", auth("admin"), async (req, res) => {
  try {
    const dept = await Department.findByIdAndUpdate(req.params.id, req.body, { new: true });
    if (req.body.hodId) await User.findByIdAndUpdate(req.body.hodId, { departmentId: dept._id });
    res.json(dept);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete("/departments/:id", auth("admin"), async (req, res) => {
  try {
    await Department.findByIdAndDelete(req.params.id);
    res.json({ message: "Deleted." });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ── Subjects ── */

app.get("/subjects/public", async (req, res) => {
  try {
    const subjects = await Subject.find({})
      .select("name code divisionName year semester type departmentId")
      .sort({ year: 1, semester: 1, name: 1 })
      .lean();
    res.json(subjects);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get("/subjects", auth(), async (req, res) => {
  try {
    const f = {};
    if (req.user.role === "hod") f.departmentId = req.user.departmentId;
    else if (req.user.role === "subject_teacher")
      f._id = { $in: (req.user.assignedSubjects || []).map(a => a.subjectId) };
    else if (req.query.departmentId) f.departmentId = req.query.departmentId;
    if (req.query.year)       f.year       = +req.query.year;
    if (req.query.semester)   f.semester   = +req.query.semester;
    if (req.query.divisionId) f.divisionId = req.query.divisionId;
    if (req.query.teacherId)  f.teacherId  = req.query.teacherId;
    const subjects = await Subject.find(f).populate("teacherId", "name email").lean();
    res.json(subjects);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/subjects", auth("admin", "hod"), async (req, res) => {
  try {
    const data = { ...req.body };
    if (req.user.role === "hod") data.departmentId = req.user.departmentId;
    const subject = await Subject.create(data);
    if (data.teacherId) {
      await User.findByIdAndUpdate(data.teacherId, {
        $addToSet: { assignedSubjects: { subjectId: subject._id, divisionId: data.divisionId } }
      });
    }
    res.json(subject);
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.patch("/subjects/:id", auth("admin", "hod"), async (req, res) => {
  try {
    const old = await Subject.findById(req.params.id);
    if (req.body.teacherId !== undefined && old.teacherId?.toString() !== req.body.teacherId) {
      if (old.teacherId)
        await User.findByIdAndUpdate(old.teacherId, { $pull: { assignedSubjects: { subjectId: old._id } } });
      if (req.body.teacherId)
        await User.findByIdAndUpdate(req.body.teacherId, {
          $addToSet: { assignedSubjects: { subjectId: old._id, divisionId: req.body.divisionId || old.divisionId } }
        });
    }
    const subject = await Subject.findByIdAndUpdate(req.params.id, req.body, { new: true });
    res.json(subject);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete("/subjects/:id", auth("admin", "hod"), async (req, res) => {
  try {
    const s = await Subject.findByIdAndDelete(req.params.id);
    if (s?.teacherId)
      await User.findByIdAndUpdate(s.teacherId, { $pull: { assignedSubjects: { subjectId: s._id } } });
    res.json({ message: "Deleted." });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ── Attendance ── */

app.post("/attendance/checkin", async (req, res) => {
  try {
    const { studentRoll, studentName, subjectId, date, lectureType } = req.body;
    if (!studentRoll || !studentName || !subjectId)
      return res.status(400).json({ error: "Missing fields." });

    const s = await Subject.findById(subjectId).lean();
    if (!s) return res.status(404).json({ error: "Subject not found." });

    let effectiveLectureType;
    if (s.type === "TH")      effectiveLectureType = "TH";
    else if (s.type === "PR") effectiveLectureType = "PR";
    else effectiveLectureType = lectureType === "PR" ? "PR" : "TH";

    const today = date || new Date().toISOString().split("T")[0];

    await Attendance.updateOne(
      { studentRoll, subjectId, date: today, lectureType: effectiveLectureType },
      { $set: {
          studentName,
          divisionId:   s.divisionId,
          departmentId: s.departmentId,
          year:         s.year,
          semester:     s.semester,
          lectureType:  effectiveLectureType
        }
      },
      { upsert: true }
    );
    res.json({ message: "Attendance recorded.", lectureType: effectiveLectureType });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/attendance/bulk", auth("admin", "hod", "subject_teacher"), async (req, res) => {
  try {
    const { subjectId, divisionId, date, records, lectureType } = req.body;
    const subject = await Subject.findById(subjectId).lean();
    if (!subject) return res.status(404).json({ error: "Subject not found." });

    let effectiveLectureType;
    if (subject.type === "TH")      effectiveLectureType = "TH";
    else if (subject.type === "PR") effectiveLectureType = "PR";
    else {
      if (!lectureType || !["TH", "PR"].includes(lectureType))
        return res.status(400).json({ error: "lectureType ('TH' or 'PR') is required for TH+PR subjects." });
      effectiveLectureType = lectureType;
    }

    const today = date || new Date().toISOString().split("T")[0];
    const present     = records.filter(r => r.present);
    const absentRolls = records.filter(r => !r.present).map(r => r.studentRoll);

    if (absentRolls.length)
      await Attendance.deleteMany({
        studentRoll: { $in: absentRolls },
        subjectId,
        date: today,
        lectureType: effectiveLectureType
      });

    if (present.length) {
      await Attendance.bulkWrite(present.map(r => ({
        updateOne: {
          filter: { studentRoll: r.studentRoll, subjectId, date: today, lectureType: effectiveLectureType },
          update: { $set: {
            studentName:  r.studentName,
            divisionId,
            departmentId: subject.departmentId,
            year:         subject.year,
            semester:     subject.semester,
            lectureType:  effectiveLectureType,
            markedBy:     req.user.id
          }},
          upsert: true
        }
      })));
    }

    res.json({ message: `${present.length} present, ${absentRolls.length} absent.`, lectureType: effectiveLectureType });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get("/attendance", auth(), async (req, res) => {
  try {
    const { subjectId, divisionId, year, semester, departmentId, range, date, from, to, lectureType } = req.query;
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
    if (lectureType && ["TH", "PR"].includes(lectureType)) f.lectureType = lectureType;

    if (date) f.date = date;
    else if (from || to) f.date = { ...(from && { $gte: from }), ...(to && { $lte: to }) };
    else if (range) { const r = getDateRange(range); f.date = { $gte: r.from, $lte: r.to }; }

    res.json(await Attendance.find(f).sort({ date: 1 }).lean());
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.patch("/attendance", auth("admin", "hod", "subject_teacher"), async (req, res) => {
  try {
    const { studentRoll, subjectId, date, present, studentName, lectureType } = req.body;

    const s = await Subject.findById(subjectId).lean();
    if (!s) return res.status(404).json({ error: "Subject not found." });

    let effectiveLectureType;
    if (s.type === "TH")      effectiveLectureType = "TH";
    else if (s.type === "PR") effectiveLectureType = "PR";
    else {
      if (!lectureType || !["TH", "PR"].includes(lectureType))
        return res.status(400).json({ error: "lectureType ('TH' or 'PR') is required for TH+PR subjects." });
      effectiveLectureType = lectureType;
    }

    if (present) {
      await Attendance.updateOne(
        { studentRoll, subjectId, date, lectureType: effectiveLectureType },
        { $set: {
            studentName,
            divisionId:   s.divisionId,
            departmentId: s.departmentId,
            year:         s.year,
            semester:     s.semester,
            lectureType:  effectiveLectureType,
            markedBy:     req.user.id
          }
        },
        { upsert: true }
      );
    } else {
      await Attendance.deleteOne({ studentRoll, subjectId, date, lectureType: effectiveLectureType });
    }
    res.json({ message: "Updated.", lectureType: effectiveLectureType });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

/* ── Export Excel ── */

app.get("/export/:subjectId", auth(), async (req, res) => {
  try {
    const subject = await Subject.findById(req.params.subjectId)
      .populate("departmentId", "name code years").lean();
    if (!subject) return res.status(404).json({ error: "Subject not found." });

    const { range, from, to, lectureType: filterType } = req.query;
    const dateF = range
      ? (() => { const r = getDateRange(range); return { $gte: r.from, $lte: r.to }; })()
      : (from || to) ? { ...(from && { $gte: from }), ...(to && { $lte: to }) } : null;

    const baseQuery = {
      subjectId: req.params.subjectId,
      ...(dateF && { date: dateF })
    };

    let typesToExport;
    if (subject.type === "TH")      typesToExport = ["TH"];
    else if (subject.type === "PR") typesToExport = ["PR"];
    else if (filterType && ["TH", "PR"].includes(filterType)) typesToExport = [filterType];
    else                            typesToExport = ["TH", "PR"];

    const wb = new ExcelJS.Workbook();

    async function buildSheet(lt) {
      const records = await Attendance.find({ ...baseQuery, lectureType: lt }).sort({ date: 1 }).lean();
      if (!records.length) return null;

      const dates    = [...new Set(records.map(r => r.date))].sort();
      const students = {};
      records.forEach(r => {
        students[r.studentRoll] ??= { roll: r.studentRoll, name: r.studentName, att: {} };
        students[r.studentRoll].att[r.date] = true;
      });

      const label = lt === "TH" ? "Theory" : "Practical";
      const ws    = wb.addWorksheet(`${subject.code}-${subject.divisionName}-${label}`);

      const headerRow = ws.addRow(["Roll", "Name", ...dates.map(fmtDate), "Present", "Total", "%"]);
      headerRow.font = { bold: true };
      headerRow.fill = {
        type: "pattern", pattern: "solid",
        fgColor: { argb: lt === "TH" ? "FFD6E4FF" : "FFFFD6D6" }
      };

      const sortedStudents = Object.values(students).sort((a, b) => +a.roll - +b.roll);
      sortedStudents.forEach(s => {
        const attended = dates.filter(d => s.att[d]).length;
        ws.addRow([
          s.roll, s.name,
          ...dates.map(d => s.att[d] ? "P" : "A"),
          attended, dates.length,
          dates.length ? ((attended / dates.length) * 100).toFixed(1) + "%" : "N/A"
        ]);
      });

      return { students: sortedStudents, dates, label };
    }

    const sheetResults = {};
    for (const lt of typesToExport) {
      sheetResults[lt] = await buildSheet(lt);
    }

    const hasAnyData = Object.values(sheetResults).some(r => r !== null);
    if (!hasAnyData) return res.status(404).send("No data.");

    if (typesToExport.length === 2 && sheetResults["TH"] && sheetResults["PR"]) {
      const summaryWs = wb.addWorksheet("Summary");
      const sumHeader = summaryWs.addRow(["Roll", "Name", "TH Present", "TH Total", "TH %", "PR Present", "PR Total", "PR %", "Overall %"]);
      sumHeader.font = { bold: true };
      sumHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2F0D9" } };

      const allRolls = new Set([
        ...sheetResults["TH"].students.map(s => s.roll),
        ...sheetResults["PR"].students.map(s => s.roll)
      ]);

      const thMap   = Object.fromEntries(sheetResults["TH"].students.map(s => [s.roll, s]));
      const prMap   = Object.fromEntries(sheetResults["PR"].students.map(s => [s.roll, s]));
      const thTotal = sheetResults["TH"].dates.length;
      const prTotal = sheetResults["PR"].dates.length;

      [...allRolls].sort((a, b) => +a - +b).forEach(roll => {
        const thStudent = thMap[roll];
        const prStudent = prMap[roll];
        const name      = thStudent?.name || prStudent?.name || "";

        const thPresent = thStudent ? sheetResults["TH"].dates.filter(d => thStudent.att[d]).length : 0;
        const prPresent = prStudent ? sheetResults["PR"].dates.filter(d => prStudent.att[d]).length : 0;

        const thPct  = thTotal ? ((thPresent / thTotal) * 100).toFixed(1) + "%" : "N/A";
        const prPct  = prTotal ? ((prPresent / prTotal) * 100).toFixed(1) + "%" : "N/A";
        const allPct = (thTotal + prTotal)
          ? (((thPresent + prPresent) / (thTotal + prTotal)) * 100).toFixed(1) + "%"
          : "N/A";

        summaryWs.addRow([roll, name, thPresent, thTotal, thPct, prPresent, prTotal, prPct, allPct]);
      });
    }

    res.setHeader("Content-Disposition",
      `attachment; filename=${subject.code}_${subject.divisionName}_Y${subject.year}S${subject.semester}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error(e);
    res.status(500).send("Export failed.");
  }
});

/* ── Health & Fallback ── */

app.get("/health", (_, res) => res.json({ ok: true }));
app.get("/", (req, res) => res.sendFile(path.join(__dirname, "public", "admin.html")));
app.use((_, res) => res.status(404).send("Not found."));

/* ── Start (local dev only) ── */

if (process.env.NODE_ENV !== "production") {
  app.listen(process.env.PORT || 3000, () => console.log("Server running on port 3000."));
}

module.exports = app;