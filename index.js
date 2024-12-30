const express = require("express");
const app = express();
const product = require("./api/product");
const chart = require("./api/chart")

app.use(express.json({ extended: false }));

app.use("/api/product", product);
app.use("/api/chart", chart);
app.use("/api/chartbase64", chart);

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`Server is running in port ${PORT}`));
