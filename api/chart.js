const express = require("express");
const {createCanvas, registerFont} = require("canvas");
const echarts = require("echarts");
const {join} = require("path");
const router = express.Router();

// 加载中文字体
registerFont(join(__dirname, '../fonts/NotoSansCJK-Regular.ttc'), { family: 'Noto Sans CJK' });

/**
 * GET product list.
 *
 * @return product list | empty.
 */
router.get('/', (req, res) => {
  const width = 800;
  const height = 600;
  const canvas = createCanvas(width, height);
  const chart = echarts.init(canvas);

  const option = {
    title: {
      text: 'Basic Radar Chart'
    },
    legend: {
      data: ['张晓名', '平均分']
    },
    radar: {
      indicator: [
        { name: '爱国情怀', max: 6500 },
        { name: '正直诚信', max: 16000 },
        { name: '遵规守纪', max: 30000 },
        { name: '仁爱友善', max: 38000 },
        { name: '包容理解', max: 52000 }
      ]
    },
    series: [
      {
        name: 'Budget vs spending',
        type: 'radar',
        data: [
          {
            value: [4200, 3000, 20000, 35000, 50000],
            name: '张晓名'
          },
          {
            value: [5000, 14000, 28000, 26000, 42000],
            name: '平均分'
          }
        ]
      }
    ]
  };

  chart.setOption(option);

  const buffer = canvas.toBuffer('image/png');
  res.set('Content-Type', 'image/png');
  res.send(buffer);
});

module.exports = router;
