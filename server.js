const express=require('express')
const cors=require('cors');
const data  = require('./src/data/data-20C-5KR');
const app = express();
const port = 3010;
app.use(cors())
app.get("/data",(req,res)=>{
    res.send(data)
})

app.listen(port,()=>{
    console.log("Server is listening")
})