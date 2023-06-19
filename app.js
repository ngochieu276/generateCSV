const express = require("express");
const app = express();

const { exportCSV } = require("./export");
app.listen(3000, () => {
  console.log("Application started and Listening on port 3001");
});

app.get("/", (req, res) => {
  console.log("asdf");
  exportCSV();
});
