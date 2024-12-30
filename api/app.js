const express = require("express");

// 创建Express应用
const app = express();


// 定义一下假数据，用于验证接口
const users = [
  { id: 1, name: "刘玄德", role: "大哥" },
  { id: 2, name: "关云长", role: "二哥" },
  { id: 3, name: "张翼德", role: "三弟" }
];

// 简单写一个接口
app.get("/1", (req, res) => {
  res.send("这是一个Node express简单服务。");
});


// 简写写一个获取用户的接口
app.get("/user", (req, res) => {
  res.status(200).json({
    code: 200,
    msg: "ok",
    data: users
  });
});


// 运行服务器
app.listen(9000, () => {
  console.log("Express Server running at http://127.0.0.1:9000");
});

/**
 * 为了让Vercel将Express转变为无服务器功能，
 * 必须导出Express应用。
 */
module.exports = app;odule.exports = app;
