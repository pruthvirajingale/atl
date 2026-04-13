const express = require("express");
const mongoose = require("mongoose");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const ExcelJS = require("exceljs");
const cors = require("cors");
const path = require("path");
const multer = require("multer");

const app = express();
const SECRET = process.env.JWT_SECRET || "change_this_secret";

// multer — keep uploaded file in memory, no disk writes
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

/* ── Database ── */

let isConnected = false;

async function connectDB() {
  if (isConnected || mongoose.connection.readyState >= 1) {
    isConnected = true;
    return;
  }
  await mongoose.connect(
    process.env.MONGODB_URI ||
      "mongodb+srv://ingalepruthviraj50_db_user:iHKXlcgpv9DLG1xS@cluster99.apwbb2y.mongodb.net/",
    {
      serverSelectionTimeoutMS: 5000,
      bufferCommands: false,
    },
  );
  isConnected = true;
}

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

const User = mongoose.model(
  "User",
  new mongoose.Schema({
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true, lowercase: true },
    password: { type: String, required: true },
    role: {
      type: String,
      enum: ["admin", "hod", "class_teacher", "subject_teacher"],
      required: true,
    },
    departmentId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Department",
      default: null,
    },
    assignedSubjects: [
      {
        subjectId: mongoose.Schema.Types.ObjectId,
        divisionId: mongoose.Schema.Types.ObjectId,
      },
    ],
  }),
);

const Department = mongoose.model(
  "Department",
  new mongoose.Schema({
    name: { type: String, required: true },
    code: { type: String, required: true, unique: true, uppercase: true },
    hodId: { type: mongoose.Schema.Types.ObjectId, ref: "User", default: null },
    years: [
      {
        year: { type: Number, enum: [1, 2, 3] },
        divisions: [
          { name: String, rollRange: { start: Number, end: Number } },
        ],
      },
    ],
  }),
);

const Subject = mongoose.model(
  "Subject",
  new mongoose.Schema({
    name: { type: String, required: true },
    code: { type: String, required: true },
    departmentId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Department",
      required: true,
    },
    year: { type: Number, enum: [1, 2, 3], required: true },
    semester: { type: Number, enum: [1, 2], required: true },
    type: { type: String, enum: ["TH", "PR", "TH+PR"], required: true },
    divisionId: { type: mongoose.Schema.Types.ObjectId, required: true },
    divisionName: { type: String, required: true },
    teacherId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      default: null,
    },
  }),
);

const attendanceSchema = new mongoose.Schema({
  studentRoll: { type: String, required: true },
  studentName: { type: String, required: true },
  subjectId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Subject",
    required: true,
  },
  departmentId: { type: mongoose.Schema.Types.ObjectId, required: true },
  divisionId: { type: mongoose.Schema.Types.ObjectId, required: true },
  year: Number,
  semester: Number,
  date: { type: String, required: true },
  lectureType: { type: String, enum: ["TH", "PR"], default: "TH" },
  markedBy: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "User",
    default: null,
  },
});
attendanceSchema.index(
  { studentRoll: 1, subjectId: 1, date: 1, lectureType: 1 },
  { unique: true },
);
const Attendance = mongoose.model("Attendance", attendanceSchema);

// Student roster — source of truth for who belongs to a division
const studentSchema = new mongoose.Schema({
  roll: { type: String, required: true },
  name: { type: String, required: true },
  departmentId: { type: mongoose.Schema.Types.ObjectId, required: true },
  divisionId: { type: mongoose.Schema.Types.ObjectId, required: true },
  year: { type: Number, required: true },
});
studentSchema.index({ roll: 1, divisionId: 1 }, { unique: true });
const Student = mongoose.model("Student", studentSchema);

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

const hashPw = (pw) => bcrypt.hash(pw, 10);
const checkPw = (pw, hash) => bcrypt.compare(pw, hash);

function makeToken(user) {
  return jwt.sign(
    {
      id: user._id,
      role: user.role,
      departmentId: user.departmentId,
      assignedSubjects: user.assignedSubjects || [],
    },
    SECRET,
    { expiresIn: "7d" },
  );
}

function getDateRange(range) {
  const to = new Date(),
    from = new Date();
  if (range === "week") from.setDate(to.getDate() - 6);
  if (range === "2weeks") from.setDate(to.getDate() - 13);
  if (range === "month") from.setMonth(to.getMonth() - 1);
  return {
    from: from.toISOString().split("T")[0],
    to: to.toISOString().split("T")[0],
  };
}

function buildDateFilter(query) {
  const { range, date, from, to } = query;
  if (date) return date;
  if (from || to)
    return { ...(from && { $gte: from }), ...(to && { $lte: to }) };
  if (range) {
    const r = getDateRange(range);
    return { $gte: r.from, $lte: r.to };
  }
  return null;
}

function resolveLectureType(subjectType, requested) {
  if (subjectType === "TH") return "TH";
  if (subjectType === "PR") return "PR";
  return requested;
}

const MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
function fmtDate(iso) {
  const [y, m, d] = iso.split("-");
  return `${d} ${MONTHS[+m - 1]} ${y.slice(2)}`;
}

/* ── Auth ── */

app.get("/auth/check-admin", async (req, res) => {
  try {
    res.json({ exists: !!(await User.exists({ role: "admin" })) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/auth/signup", async (req, res) => {
  try {
    const { name, email, password } = req.body;
    if (!name || !email || !password)
      return res.status(400).json({ error: "All fields required." });
    if (await User.findOne({ email }))
      return res.status(409).json({ error: "Email already in use." });
    const user = await User.create({
      name,
      email,
      password: await hashPw(password),
      role: "admin",
    });
    res.json({ token: makeToken(user), name: user.name, role: user.role });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/auth/login", async (req, res) => {
  try {
    const { email, password } = req.body;
    const user = await User.findOne({ email });
    if (!user || !(await checkPw(password, user.password)))
      return res.status(401).json({ error: "Invalid email or password." });
    res.json({
      token: makeToken(user),
      name: user.name,
      role: user.role,
      departmentId: user.departmentId,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ── Users ── */

app.get("/users", auth("admin", "hod"), async (req, res) => {
  try {
    const filter = {};
    if (req.query.role) filter.role = req.query.role;
    if (req.user.role === "hod") filter.departmentId = req.user.departmentId;
    else if (req.query.departmentId)
      filter.departmentId = req.query.departmentId;
    res.json(await User.find(filter, "-password").lean());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/users", auth("admin", "hod"), async (req, res) => {
  try {
    const { name, email, password, role } = req.body;
    if (!name || !email || !password || !role)
      return res.status(400).json({ error: "All fields required." });
    if (
      req.user.role === "hod" &&
      !["class_teacher", "subject_teacher"].includes(role)
    )
      return res
        .status(403)
        .json({ error: "HOD can only create teacher accounts." });
    if (await User.findOne({ email }))
      return res.status(409).json({ error: "Email in use." });
    const departmentId =
      req.user.role === "hod" ? req.user.departmentId : req.body.departmentId;
    const user = await User.create({
      name,
      email,
      password: await hashPw(password),
      role,
      departmentId,
    });
    res.json({ _id: user._id, name: user.name, role: user.role });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.patch("/users/:id", auth("admin"), async (req, res) => {
  try {
    res.json(
      await User.findByIdAndUpdate(req.params.id, req.body, {
        new: true,
        select: "-password",
      }),
    );
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete("/users/:id", auth("admin", "hod"), async (req, res) => {
  try {
    const user = await User.findById(req.params.id);
    if (!user) return res.status(404).json({ error: "User not found." });
    if (
      req.user.role === "hod" &&
      user.departmentId?.toString() !== req.user.departmentId?.toString()
    )
      return res.status(403).json({ error: "Access denied." });
    await User.findByIdAndDelete(req.params.id);
    res.json({ message: "Deleted." });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ── Departments ── */

app.get("/departments", auth(), async (req, res) => {
  try {
    res.json(await Department.find().populate("hodId", "name email").lean());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/departments/:id", auth(), async (req, res) => {
  try {
    res.json(
      await Department.findById(req.params.id)
        .populate("hodId", "name email")
        .lean(),
    );
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/departments", auth("admin"), async (req, res) => {
  try {
    const dept = await Department.create(req.body);
    if (req.body.hodId)
      await User.findByIdAndUpdate(req.body.hodId, { departmentId: dept._id });
    res.json(dept);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.patch("/departments/:id", auth("admin"), async (req, res) => {
  try {
    const dept = await Department.findByIdAndUpdate(req.params.id, req.body, {
      new: true,
    });
    if (req.body.hodId)
      await User.findByIdAndUpdate(req.body.hodId, { departmentId: dept._id });
    res.json(dept);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete("/departments/:id", auth("admin"), async (req, res) => {
  try {
    await Department.findByIdAndDelete(req.params.id);
    res.json({ message: "Deleted." });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ── Subjects ── */

app.get("/subjects/public", async (req, res) => {
  try {
    res.json(
      await Subject.find({})
        .select("name code divisionName year semester type departmentId")
        .sort({ year: 1, semester: 1, name: 1 })
        .lean(),
    );
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/subjects", auth(), async (req, res) => {
  try {
    const f = {};
    if (req.user.role === "hod") f.departmentId = req.user.departmentId;
    else if (req.user.role === "subject_teacher")
      f._id = {
        $in: (req.user.assignedSubjects || []).map((a) => a.subjectId),
      };
    else if (req.query.departmentId) f.departmentId = req.query.departmentId;

    if (req.query.year) f.year = +req.query.year;
    if (req.query.semester) f.semester = +req.query.semester;
    if (req.query.divisionId)
      f.divisionId = new mongoose.Types.ObjectId(req.query.divisionId);
    if (req.query.teacherId) f.teacherId = req.query.teacherId;

    res.json(await Subject.find(f).populate("teacherId", "name email").lean());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/subjects", auth("admin", "hod"), async (req, res) => {
  try {
    const data =
      req.user.role === "hod"
        ? { ...req.body, departmentId: req.user.departmentId }
        : { ...req.body };
    const subject = await Subject.create(data);
    if (data.teacherId)
      await User.findByIdAndUpdate(data.teacherId, {
        $addToSet: {
          assignedSubjects: {
            subjectId: subject._id,
            divisionId: data.divisionId,
          },
        },
      });
    res.json(subject);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.patch("/subjects/:id", auth("admin", "hod"), async (req, res) => {
  try {
    const old = await Subject.findById(req.params.id);
    if (
      req.body.teacherId !== undefined &&
      old.teacherId?.toString() !== req.body.teacherId
    ) {
      if (old.teacherId)
        await User.findByIdAndUpdate(old.teacherId, {
          $pull: { assignedSubjects: { subjectId: old._id } },
        });
      if (req.body.teacherId)
        await User.findByIdAndUpdate(req.body.teacherId, {
          $addToSet: {
            assignedSubjects: {
              subjectId: old._id,
              divisionId: req.body.divisionId || old.divisionId,
            },
          },
        });
    }
    res.json(
      await Subject.findByIdAndUpdate(req.params.id, req.body, { new: true }),
    );
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete("/subjects/:id", auth("admin", "hod"), async (req, res) => {
  try {
    const s = await Subject.findByIdAndDelete(req.params.id);
    if (s?.teacherId)
      await User.findByIdAndUpdate(s.teacherId, {
        $pull: { assignedSubjects: { subjectId: s._id } },
      });
    res.json({ message: "Deleted." });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ── Students ── */

// GET /students?divisionId=&year=&departmentId=
app.get("/students", auth(), async (req, res) => {
  try {
    const f = {};
    if (["hod", "class_teacher"].includes(req.user.role))
      f.departmentId = req.user.departmentId;
    else if (req.query.departmentId && req.user.role === "admin")
      f.departmentId = req.query.departmentId;

    if (req.query.divisionId)
      f.divisionId = new mongoose.Types.ObjectId(req.query.divisionId);
    if (req.query.year) f.year = +req.query.year;

    res.json(await Student.find(f).sort({ roll: 1 }).lean());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /students/bulk
// Body: [{ roll, name, divisionId, year }]
// HOD's departmentId is taken from their token; admin must pass departmentId per row.
app.post("/students/bulk", auth("admin", "hod"), async (req, res) => {
  try {
    const rows = req.body;
    if (!Array.isArray(rows) || !rows.length)
      return res.status(400).json({ error: "Send an array of students." });

    const deptId = req.user.role === "hod" ? req.user.departmentId : null;

    const ops = rows.map((s) => ({
      updateOne: {
        filter: {
          roll: String(s.roll),
          divisionId: new mongoose.Types.ObjectId(s.divisionId),
        },
        update: {
          $set: {
            name: String(s.name),
            year: +s.year,
            departmentId: deptId
              ? new mongoose.Types.ObjectId(deptId)
              : new mongoose.Types.ObjectId(s.departmentId),
            divisionId: new mongoose.Types.ObjectId(s.divisionId),
          },
        },
        upsert: true,
      },
    }));

    const result = await Student.bulkWrite(ops);
    res.json({
      upserted: result.upsertedCount,
      modified: result.modifiedCount,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /students/upload
// Multipart form-data: field "file" (.xlsx or .csv), columns: roll | name  (header row required)
// HOD picks year + divisionId in the UI — passed as query params
// HOD's departmentId is taken from their token; admin must also pass departmentId as a query param
//
// Query params:
//   ?divisionId=<ObjectId>   (required)
//   ?year=<1|2|3>            (required)
//   ?departmentId=<ObjectId> (required for admin only)
//
// Excel / CSV format (first row = headers):
//   roll  | name
//   101   | Priya Sharma
//   102   | Rahul Mehta

app.post(
  "/students/upload",
  auth("admin", "hod"),
  upload.single("file"),
  async (req, res) => {
    try {
      if (!req.file)
        return res
          .status(400)
          .json({
            error:
              "No file uploaded. Send the Excel/CSV as form-data field 'file'.",
          });

      const { divisionId, year } = req.query;

      if (!divisionId || !year)
        return res
          .status(400)
          .json({
            error: "Query params 'divisionId' and 'year' are required.",
          });

      if (!mongoose.Types.ObjectId.isValid(divisionId))
        return res.status(400).json({ error: "Invalid divisionId." });

      // departmentId — from token for HOD, from query for admin
      const departmentId =
        req.user.role === "hod"
          ? req.user.departmentId
          : req.query.departmentId;

      if (!departmentId)
        return res
          .status(400)
          .json({
            error: "Admin must provide 'departmentId' as a query param.",
          });

      if (!mongoose.Types.ObjectId.isValid(departmentId))
        return res.status(400).json({ error: "Invalid departmentId." });

      // ── Load workbook ──
      const wb = new ExcelJS.Workbook();
      const ext = req.file.originalname.split(".").pop().toLowerCase();

      if (["xlsx", "xls"].includes(ext)) {
        await wb.xlsx.load(req.file.buffer);
      } else if (ext === "csv") {
        const { Readable } = require("stream");
        await wb.csv.read(Readable.from(req.file.buffer));
      } else {
        return res
          .status(400)
          .json({ error: "Only .xlsx or .csv files are supported." });
      }

      const ws = wb.worksheets[0];
      if (!ws)
        return res
          .status(400)
          .json({ error: "File appears to be empty or unreadable." });

      // ── Detect header columns (case-insensitive) ──
      const headerValues = ws.getRow(1).values; // index 0 is always undefined in ExcelJS
      const colIndex = {};
      headerValues.forEach((h, i) => {
        if (h) colIndex[String(h).trim().toLowerCase()] = i;
      });

      if (colIndex["roll"] === undefined || colIndex["name"] === undefined) {
        return res.status(400).json({
          error:
            `First row must contain column headers "roll" and "name". ` +
            `Found: ${
              headerValues
                .filter(Boolean)
                .map((h) => `"${h}"`)
                .join(", ") || "nothing"
            }.`,
        });
      }

      // ── Parse rows ──
      const rows = [];
      const errors = [];

      ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return; // skip header

        const roll = String(row.values[colIndex["roll"]] ?? "").trim();
        const name = String(row.values[colIndex["name"]] ?? "").trim();

        // silently skip completely blank rows
        if (!roll && !name) return;

        if (!roll) {
          errors.push(`Row ${rowNum}: missing roll number (name: "${name}").`);
          return;
        }
        if (!name) {
          errors.push(`Row ${rowNum}: missing name (roll: "${roll}").`);
          return;
        }
        if (isNaN(+roll)) {
          errors.push(`Row ${rowNum}: roll "${roll}" must be numeric.`);
          return;
        }

        rows.push({ roll, name });
      });

      if (!rows.length)
        return res
          .status(400)
          .json({
            error: "No valid student rows found in the file.",
            details: errors,
          });

      // ── Warn about duplicate rolls within the uploaded file ──
      const seen = new Set();
      rows.forEach((r) => {
        if (seen.has(r.roll))
          errors.push(
            `Duplicate roll ${r.roll} in uploaded file — only the last occurrence will be saved.`,
          );
        seen.add(r.roll);
      });

      // ── Upsert into DB ──
      const ops = rows.map((s) => ({
        updateOne: {
          filter: {
            roll: s.roll,
            divisionId: new mongoose.Types.ObjectId(divisionId),
          },
          update: {
            $set: {
              name: s.name,
              year: +year,
              departmentId: new mongoose.Types.ObjectId(departmentId),
              divisionId: new mongoose.Types.ObjectId(divisionId),
            },
          },
          upsert: true,
        },
      }));

      const result = await Student.bulkWrite(ops);

      res.json({
        success: true,
        upserted: result.upsertedCount, // brand-new students added
        modified: result.modifiedCount, // existing students whose name was updated
        total: rows.length,
        errors, // row-level issues the HOD can fix and re-upload
      });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

/* ── Attendance ── */

app.post("/attendance/checkin", async (req, res) => {
  try {
    const { studentRoll, studentName, subjectId, date, lectureType } = req.body;
    if (!studentRoll || !studentName || !subjectId)
      return res.status(400).json({ error: "Missing fields." });

    const s = await Subject.findById(subjectId).lean();
    if (!s) return res.status(404).json({ error: "Subject not found." });

    // Division validation
    const dept = await Department.findById(s.departmentId).lean();
    if (dept) {
      const yearObj = (dept.years || []).find((y) => y.year === s.year);
      const subjectDiv = (yearObj?.divisions || []).find(
        (d) => d._id.toString() === s.divisionId.toString(),
      );
      if (subjectDiv?.rollRange) {
        const roll = parseInt(studentRoll, 10);
        if (isNaN(roll))
          return res
            .status(403)
            .json({ error: `Roll number "${studentRoll}" must be numeric.` });
        const studentDiv = (yearObj.divisions || []).find(
          (d) => roll >= d.rollRange.start && roll <= d.rollRange.end,
        );
        if (!studentDiv)
          return res
            .status(403)
            .json({
              error: `Roll ${studentRoll} does not belong to any division in Year ${s.year}.`,
            });
        if (studentDiv._id.toString() !== s.divisionId.toString())
          return res.status(403).json({
            error: `Roll ${studentRoll} belongs to Division ${studentDiv.name}, not Division ${subjectDiv.name}.`,
          });
      }
    }

    const effectiveLectureType = resolveLectureType(
      s.type,
      lectureType === "PR" ? "PR" : "TH",
    );
    const today = date || new Date().toISOString().split("T")[0];

    // Record attendance
    await Attendance.updateOne(
      {
        studentRoll,
        subjectId,
        date: today,
        lectureType: effectiveLectureType,
      },
      {
        $set: {
          studentName,
          divisionId: s.divisionId,
          departmentId: s.departmentId,
          year: s.year,
          semester: s.semester,
          lectureType: effectiveLectureType,
        },
      },
      { upsert: true },
    );

    // Update roster — $set (not $setOnInsert) so QR scans refresh names of pre-seeded students
    await Student.updateOne(
      { roll: studentRoll, divisionId: s.divisionId },
      {
        $set: {
          name: studentName,
          departmentId: s.departmentId,
          year: s.year,
        },
      },
      { upsert: true },
    );

    res.json({
      message: "Attendance recorded.",
      lectureType: effectiveLectureType,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post(
  "/attendance/bulk",
  auth("admin", "hod", "subject_teacher"),
  async (req, res) => {
    try {
      const { subjectId, date, records, lectureType } = req.body;
      const subject = await Subject.findById(subjectId).lean();
      if (!subject)
        return res.status(404).json({ error: "Subject not found." });

      const effectiveLectureType = resolveLectureType(
        subject.type,
        lectureType,
      );
      if (!effectiveLectureType || !["TH", "PR"].includes(effectiveLectureType))
        return res
          .status(400)
          .json({
            error: "lectureType ('TH' or 'PR') is required for TH+PR subjects.",
          });

      const today = date || new Date().toISOString().split("T")[0];
      const present = records.filter((r) => r.present);
      const absentRolls = records
        .filter((r) => !r.present)
        .map((r) => r.studentRoll);

      if (absentRolls.length)
        await Attendance.deleteMany({
          studentRoll: { $in: absentRolls },
          subjectId,
          date: today,
          lectureType: effectiveLectureType,
        });

      if (present.length) {
        await Attendance.bulkWrite(
          present.map((r) => ({
            updateOne: {
              filter: {
                studentRoll: r.studentRoll,
                subjectId,
                date: today,
                lectureType: effectiveLectureType,
              },
              update: {
                $set: {
                  studentName: r.studentName,
                  divisionId: subject.divisionId,
                  departmentId: subject.departmentId,
                  year: subject.year,
                  semester: subject.semester,
                  lectureType: effectiveLectureType,
                  markedBy: req.user.id,
                },
              },
              upsert: true,
            },
          })),
        );
      }

      res.json({
        message: `${present.length} present, ${absentRolls.length} absent.`,
        lectureType: effectiveLectureType,
      });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

app.get("/attendance", auth(), async (req, res) => {
  try {
    const { subjectId, divisionId, year, semester, departmentId, lectureType } =
      req.query;
    const f = {};

    if (req.user.role === "subject_teacher")
      f.subjectId = {
        $in: (req.user.assignedSubjects || []).map((a) => a.subjectId),
      };
    else if (["class_teacher", "hod"].includes(req.user.role))
      f.departmentId = req.user.departmentId;

    if (subjectId) f.subjectId = subjectId;
    if (divisionId) f.divisionId = new mongoose.Types.ObjectId(divisionId);
    if (year) f.year = +year;
    if (semester) f.semester = +semester;
    if (departmentId && req.user.role === "admin")
      f.departmentId = departmentId;
    if (lectureType && ["TH", "PR"].includes(lectureType))
      f.lectureType = lectureType;

    const dateF = buildDateFilter(req.query);
    if (dateF) f.date = dateF;

    res.json(await Attendance.find(f).sort({ date: 1 }).lean());
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.patch(
  "/attendance",
  auth("admin", "hod", "subject_teacher"),
  async (req, res) => {
    try {
      const {
        studentRoll,
        subjectId,
        date,
        present,
        studentName,
        lectureType,
      } = req.body;

      const s = await Subject.findById(subjectId).lean();
      if (!s) return res.status(404).json({ error: "Subject not found." });

      const effectiveLectureType = resolveLectureType(s.type, lectureType);
      if (!effectiveLectureType || !["TH", "PR"].includes(effectiveLectureType))
        return res
          .status(400)
          .json({
            error: "lectureType ('TH' or 'PR') is required for TH+PR subjects.",
          });

      if (present) {
        await Attendance.updateOne(
          { studentRoll, subjectId, date, lectureType: effectiveLectureType },
          {
            $set: {
              studentName,
              divisionId: s.divisionId,
              departmentId: s.departmentId,
              year: s.year,
              semester: s.semester,
              lectureType: effectiveLectureType,
              markedBy: req.user.id,
            },
          },
          { upsert: true },
        );
      } else {
        await Attendance.deleteOne({
          studentRoll,
          subjectId,
          date,
          lectureType: effectiveLectureType,
        });
      }
      res.json({ message: "Updated.", lectureType: effectiveLectureType });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

/* ── Export Excel ── */

app.get("/export/:subjectId", auth(), async (req, res) => {
  try {
    const subject = await Subject.findById(req.params.subjectId)
      .populate("departmentId", "name code years")
      .lean();
    if (!subject) return res.status(404).json({ error: "Subject not found." });

    const dateF = buildDateFilter(req.query);

    const [records, roster] = await Promise.all([
      Attendance.find({
        subjectId: req.params.subjectId,
        ...(dateF && { date: dateF }),
      })
        .sort({ date: 1 })
        .lean(),
      Student.find({ divisionId: subject.divisionId }).sort({ roll: 1 }).lean(),
    ]);

    if (!records.length && !roster.length)
      return res.status(404).send("No data.");

    // Build sorted unique (date, lectureType) columns
    const colMap = new Map();
    records.forEach((r) => {
      const key = `${r.date}__${r.lectureType}`;
      if (!colMap.has(key))
        colMap.set(key, { date: r.date, lectureType: r.lectureType });
    });
    const cols = [...colMap.values()].sort((a, b) =>
      a.date !== b.date
        ? a.date.localeCompare(b.date)
        : a.lectureType.localeCompare(b.lectureType),
    );

    // Seed all roster students first (so absent-only students appear), then overlay attendance
    const studentMap = {};
    roster.forEach((s) => {
      studentMap[s.roll] = { roll: s.roll, name: s.name, att: {} };
    });
    records.forEach((r) => {
      studentMap[r.studentRoll] ??= {
        roll: r.studentRoll,
        name: r.studentName,
        att: {},
      };
      studentMap[r.studentRoll].att[`${r.date}__${r.lectureType}`] = true;
    });
    const students = Object.values(studentMap).sort(
      (a, b) => +a.roll - +b.roll,
    );

    const wb = new ExcelJS.Workbook();

    // ── Main sheet ──
    const ws = wb.addWorksheet(
      `${subject.code}-${subject.divisionName}`.slice(0, 31),
    );
    const headers = [
      "Roll",
      "Name",
      ...cols.map((c) => `${fmtDate(c.date)} (${c.lectureType})`),
      "Present",
      "Total",
      "%",
    ];

    const headerRow = ws.addRow(headers);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9E1F2" },
    };
    ws.views = [{ state: "frozen", xSplit: 2, ySplit: 1 }];

    cols.forEach((c, i) => {
      const cell = headerRow.getCell(i + 3);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: c.lectureType === "TH" ? "FFD6E4FF" : "FFEAD6FF" },
      };
      cell.alignment = { horizontal: "center" };
    });

    students.forEach((s) => {
      const attended = cols.filter(
        (c) => s.att[`${c.date}__${c.lectureType}`],
      ).length;
      const total = cols.length;
      const row = ws.addRow([
        s.roll,
        s.name,
        ...cols.map((c) => (s.att[`${c.date}__${c.lectureType}`] ? "P" : "A")),
        attended,
        total,
        total ? ((attended / total) * 100).toFixed(2) + "%" : "N/A",
      ]);
      cols.forEach((c, i) => {
        const cell = row.getCell(i + 3);
        cell.font = {
          color: {
            argb: s.att[`${c.date}__${c.lectureType}`]
              ? "FF166534"
              : "FF991B1B",
          },
          bold: true,
        };
        cell.alignment = { horizontal: "center" };
      });
      row.getCell(headers.length).alignment = { horizontal: "right" };
    });

    ws.columns.forEach((col, i) => {
      col.width = Math.min(Math.max((headers[i]?.length || 0) + 2, 8), 40);
    });

    // ── Summary sheet (TH+PR only) ──
    

    res.setHeader(
      "Content-Disposition",
      `attachment; filename=${subject.code}_${subject.divisionName}_Y${subject.year}S${subject.semester}.xlsx`,
    );
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error(e);
    res.status(500).send("Export failed.");
  }
});

/* ── Health & Fallback ── */

app.get("/health", (_, res) => res.json({ ok: true }));
app.get("/", (req, res) =>
  res.sendFile(path.join(__dirname, "public", "admin.html")),
);
app.use((_, res) => res.status(404).send("Not found."));

if (process.env.NODE_ENV !== "production")
  app.listen(process.env.PORT || 3000, () =>
    console.log("Server running on port 3000."),
  );

module.exports = app;
