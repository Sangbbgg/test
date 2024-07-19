import express from "express";
const app = express();
const port = 8000;

app.get('/', (req, res) => {
  res.send('윤과장님 안녕하세요????')
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})