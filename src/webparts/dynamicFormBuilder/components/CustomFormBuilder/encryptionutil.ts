import cryptojs from "crypto-js";
import {auth} from './ConfigURL/All_URLs';


export class dataservice{

private fpart = auth.clientId.slice(0,6);
private spart = auth.authority.slice(-6);

private parts:string[] = [this.fpart,this.spart];

private jumbledfunc(){
  const order = this.parts.map(()=>  Math.floor(Math.random() * (this.parts.length - 0) ) );
  //console.log("jumbledorder",order);
  return order;
}

private buildkey(){
    const jumborder = this.jumbledfunc();
    let key:string="";
    for (const ord of jumborder ){
       key = key.concat(this.parts[ord]);
    }
    //console.log("jumbkey",key);
    const jumbseq=jumborder.join('');
    return {key,jumbseq};
}

private keybuilder(orderstr:string){

    const orderarr = orderstr?.split("").map(Number);
    let key:string ="";
    for (const order of orderarr){
        key=key.concat(this.parts[order]);
    } //console.log("builtkey",key);
    return key;
}


public encryptjson(item:any){
    //console.log("calledencrypt",item,);
    if (typeof item === "string" && item.startsWith("U2FsdGVkX1") && /^[A-Za-z0-9+/=]+$/.test(item.replace(/["']/g, "")) ) {
      return item;
    }
    const jsonstring = JSON.stringify(item);
    //console.log("stringified",jsonstring)


if (typeof item === "string" && item.startsWith("U2FsdGVkX1")) {
  return item;
}

// if(!jsonstring || jsonstring?.startsWith("U2FsdGVkX1")) return item;

const {key,jumbseq} = this.buildkey();
//console.log("keybuilt",key)
const encryptedjson = cryptojs.AES.encrypt(jsonstring,key).toString();
const finalenc = encryptedjson.concat(jumbseq);
//console.log("encryptvalue",finalenc);
return finalenc;
// return encryptedjson;

// const response = await fetch("https://smartoffice.cloudangles.com/encrypt", {
//     method: "POST",
//     headers: { "Content-Type": "application/json" },
//     body: JSON.stringify({ data:item }),
//   });

//   const result = await response.json();
//   console.log("Encrypted::", result.encryptedData,result);
//       return result.encryptedData;

  }
  

  public decryptjson(item:any){

//console.log("recieveddecryptitem",item)
if (typeof item === "string" && item.startsWith('"')) {
    item = JSON.parse(item);
  } 
  //  console.log("about to decrypt",item)
     if(!item || typeof item != "string" || !item?.startsWith("U2FsdGVkX1")) return item;

     const partslength = this.parts?.length;
     const encryptedstr:string = item.slice(0,-partslength);
     const key = this.keybuilder(item.slice(-partslength))
     //console.log("key in decrypt",key)
     const decryptedjson = cryptojs.AES.decrypt(encryptedstr,key).toString(cryptojs.enc.Utf8);
// console.log("decryptedjson",decryptedjson);
//console.log("parsedjson",JSON.parse(decryptedjson))
    //  return JSON.parse(decryptedjson);
    return decryptedjson;

    
  }

}