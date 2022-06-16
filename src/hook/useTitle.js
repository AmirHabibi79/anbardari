import { useEffect, useState } from "react";

export default function useTitle(title) {
  const [value,setValue]=useState(title)
  useEffect(()=>{
    document.title=`anbardari - ${value}`
  },[value])
  return setValue
}
