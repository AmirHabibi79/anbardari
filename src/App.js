import {HotTable} from "@handsontable/react"
import { registerAllModules } from 'handsontable/registry';
import "handsontable/dist/handsontable.min.css"
import './App.css';
import { useCallback, useEffect, useRef, useState } from "react";
import useTitle from "./hook/useTitle";
// make a sheet based on screen size for first time  
const COL_WIDTH=50 // the width of each column
const ROW_HEIGHT=23 // the height of each row
const START_COLS=Math.trunc(window.innerWidth / COL_WIDTH) // amount of columns
const START_ROWS=Math.trunc(window.innerHeight / ROW_HEIGHT) // amount of rows
function App() {
  registerAllModules()
  const hottableRef=useRef()
  const importfileRef=useRef()
  const fileNameRef=useRef("untitled")
  const setTitle=useTitle(fileNameRef.current)
  const [FAS,setFAS]=useState(false)
  const readFile=(e)=>{
    return new Promise((resolve,reject)=>{
      const f= e.target.files[0]
      fileNameRef.current=f.name
      setTitle(f.name.split(".")[0])
    if(f){
      const fr=new FileReader()
      fr.onload=data=>{
        resolve(data.target.result)
      }
      fr.readAsBinaryString(f)
    }else{
      console.log("cant read the file")
    }
    })
  }
  const toExcel=(data)=>{
    const {read}=require("xlsx")
    const sheet=read(data,{type:"binary"})
    console.log(sheet.Sheets)
    return sheet.Sheets[sheet.SheetNames[0]]
  }
  const formatExcel=(workbook)=> {
    const {utils}=require("xlsx")
    var result = {};
    var roa = utils.sheet_to_json(workbook, {
        header: 1
    });
    if (roa.length) result = roa;
    
    return result
  };
  // add row or colunm if it has less than height or width of screen
  const addRowOrColunm=(sheet)=>{
    //add colunm
    if(sheet[0].length<START_COLS){
      const amount=START_COLS - sheet[0].length
      const colunms=Array(amount).fill("")
      for(let i=0;i<sheet.length;i++){
        sheet[i].push(...colunms)
      }
    }
    //add row
    else if(sheet.length<START_ROWS){
      const amount=(START_ROWS+10) - sheet.length
      const colunms=Array(START_COLS).fill("")
      const rows=Array(amount).fill(colunms)
      sheet.push(...rows)
    }
    return sheet
  }
  const onChangeInput=useCallback(async(e)=>{
    const data=await readFile(e)
    const tx=toExcel(data)
    const formated=formatExcel(tx)
    const finalFormat=addRowOrColunm(formated)
    hottableRef.current.hotInstance.loadData(finalFormat)
  },[]) 
  const exportSheet=()=>{
    if(fileNameRef.current==="untitled"){
      const name=getUserFileName()
      hottableRef.current.hotInstance.getPlugin("ExportFile").downloadFile("csv",{filename:name})
      setTitle(name)
    }else{
      const name=fileNameRef.current.split(".")[0]

      hottableRef.current.hotInstance.getPlugin("ExportFile").downloadFile("csv",{filename:name})
    }
  }
  const increaseNumByDraggign=(start,end,data)=>{
    if(data.from.col!==data.to.col) return 
        
        let number,position
        for(let i=0;i<start.length;i++){
          if(start[i][0]===null){
            continue
          }else{
            number=start[i][0]
            position=i
            break
          }

        }
       
      if(number===undefined||position===undefined||isNaN(number))return
        let startPosition=end.from.row
        const prop=end.from.col
        const amount=(data.to.row -startPosition)+1
        const finaldata=[]
        let preNumber=parseFloat(number)
        let offset=position
        for(let j=0;j<amount;j++){
          if(j < offset){
            const arrayData=[
            startPosition,
            prop,
            ""
            ]
            startPosition++
            
            finaldata.push(arrayData)
          }
          else{
            const arrayData=[
            startPosition,
            prop,
            preNumber
            ]
            preNumber++
            startPosition++
            offset=(position + j)+1
            finaldata.push(arrayData)
          }
          
        }
        hottableRef.current.hotInstance.setDataAtRowProp(finaldata)
  }
  useEffect(()=>{
    hottableRef.current.hotInstance.addHook('afterAutofill',increaseNumByDraggign)
    importfileRef.current.addEventListener("change",onChangeInput)
    hottableRef.current.hotInstance.getPlugin("Autofill").autoInsertRow="vertical"
    
  },[onChangeInput])
  const addRow=()=>{
    hottableRef.current.hotInstance.alter("insert_row")
  }
  const getUserFileName=()=>{
    const name=prompt("enter file name","")
    return name
  }
  return (
    <>
    <label>read from file</label>
    <input ref={importfileRef} accept=".csv,.tsv,.xlsx" type="file"/>
    <button onClick={addRow}>add row</button>
    <button onClick={exportSheet} >export</button>
    <button onClick={()=>setFAS(!FAS)} >filter and search</button>
    <HotTable ref={hottableRef} filters={true} search={true}  columns={{type:"text"}} dropdownMenu={FAS} contextMenu startCols={START_COLS} startRows={START_ROWS}  colHeaders  rowHeaders={true} licenseKey='non-commercial-and-evaluation'>
    </HotTable>
    </>
  )
}

export default App;
