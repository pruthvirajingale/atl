const mongoose = require("mongoose");

// Connect to MongoDB (no options needed in Mongoose 7+)
mongoose.connect("mongodb://127.0.0.1:27017/attendanceDB")
  .then(() => console.log("MongoDB connected"))
  .catch(err => console.error(err));

const attendanceSchema = new mongoose.Schema({
  roll: String,
  name: String,
  date: String,
  subject: String,
  status: String
});

const Attendance = mongoose.model("Attendance", attendanceSchema);

function getLastWeekDates() {
  const today = new Date();
  const dates = [];
  for (let i = 1; i <= 7; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() - i);
    dates.push(d.toISOString().split("T")[0]);
  }
  return dates.reverse();
}

async function insertAttendance() {
  const student = {
    roll: "43",
    name: "Pruthviraj Ajay Ingale"
  };
  const thStatus = "Present";
  const prStatus = "Present";
  const subjectBase = "NIS";

  const dates = getLastWeekDates();
  const records = [];

  for (let i = 0; i < dates.length; i++) {
    const date = dates[i];
    const isTH = i % 2 === 0;
    const subject = isTH ? `${subjectBase}_TH` : `${subjectBase}_PR`;
    const status = isTH ? thStatus : prStatus;

    records.push({
      roll: student.roll,
      name: student.name,
      date,
      subject,
      status
    });
  }

  await Attendance.insertMany(records);
  console.log("Last week attendance inserted!");
  mongoose.connection.close();
}

insertAttendance();
