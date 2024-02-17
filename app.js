var fs = require('fs');
var fetch = require('node-fetch');
var path = require("path");
var xlsx = require("xlsx");
const params = require("./paramsList.json")
const _undrscr = require('underscore');

const filePath = path.resolve(__dirname, "Weekly Asset Report.xlsx");
const workbook = xlsx.readFile(filePath, {cellDates: true});


fetchReport('purchaseorders','PO');
fetchReport('shipments','Shipment');
fetchReport('invoices','Invoice');
fetchReport('goodsreceipts','GR');
fetchReport('payments','Payment');

async function fetchReport(listAPIName,assetName){
try{
    let headers = { "Authorization": "Bearer " + params.token }
    const method = 'GET'   
    let page=1;
    var responselist=[];

    let listAPIURL = params.baseURL.PROD + "/api/"+listAPIName+"?createdStartDate=" + params.startDate + "&createdEndDate=" + params.endDate+"&page="+page;
    const responseListAPI = await fetch(listAPIURL, { method, headers }).then(res => res.json());    //logical variable name
    responselist=responselist.concat(responseListAPI.data.result);

    while(page<responseListAPI.metadata.totalPages){   //if response has more than 1000 data then this will execute
        page++;
        listAPIURL=listAPIURL.substring(0, listAPIURL.length-1)+page;
        let responseNxtPage=await fetch(listAPIURL, { method, headers }).then(res => res.json());
        responselist=responselist.concat(responseNxtPage.data.result)
    }

    //T2 Validation data to be inserted in sheet
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets[assetName], {raw: true, cellDates: true}); //, dateNF:'mm/dd/yyyy'
    //const arrCount = xlsx.utils.sheet_to_json(workbook.Sheets["Total Count"]);
//#region PO
    let T2po=SIpo=LOANpo=MISCBUYpo=0;
    if(assetName=='PO'){
    for(let i=0;i<responselist.length;i++){   //looping through fetched JSON result            
            if(!_undrscr.isEmpty(responselist[i])){       //checking if purchase order response list is valid and not empty
                //let createdDate=new Date(responselistPO[i].createdOn);
                let orderObj=_undrscr.filter(responselist[i].ref,function(obj){ 
                    if(obj.idQualf=='ZF'){ 
                        return obj; 
                    }
                });
                switch(orderObj[0].id){
                    case "SI": SIpo++;
                    break;
                    case "T2": T2po++;
                    break;
                    case "LOAN": LOANpo++;
                    break;
                    case "MISCBUY": MISCBUYpo++;
                    break;
                    default: break;
                }
                worksheet[i]={
                        "Purchase Order Number" : responselist[i].orderNumber,
                        "Type" : orderObj[0].id,
                        "Creation Time (UTC)" : responselist[i].createdOn
                        };
            }
    }
}
    //#endregion
    
//#region Shipment
    if(assetName=='Shipment'){
    for(let i=0;i<responselist.length;i++){   //looping through fetched JSON result            
        if(!_undrscr.isEmpty(responselist[i])){       //checking if Shipment response list is valid and not empty

            worksheet[i]={
                    "Shipment Number" : responselist[i].shipmentNumber,
                    "Creation Time (UTC)" : responselist[i].createdOn
                    };
            }
        }
    }
//#endregion

//#region Invoice
    let invLOAN=invSETTLEMENT=loanCR=settlementCR=invSI=invT2=invMISCBUY=siCR=t2CR=miscbuyCR=0;
    if(assetName=='Invoice'){
    for(let i=0;i<responselist.length;i++){   //looping through fetched JSON result            
        if(!_undrscr.isEmpty(responselist[i])){       //checking if Invoice response list is valid and not empty
          
            let invObj=_undrscr.filter(responselist[i].ref,function(obj){ 
                if(obj.idQualf=='ZF'){
                    return obj; 
                }
            });
            switch(invObj[0].id){
                case "SI": if(responselist[i].invType=="CR"){
                                siCR++;
                            }
                            else if(responselist[i].invType=="DI"){
                                invSI++;
                            }
                break;
                case "T2": if(responselist[i].invType=="CR"){
                                t2CR++;
                            }
                            else if(responselist[i].invType=="DI"){
                                invT2++;
                            }
                break;
                case "LOAN": if(responselist[i].invType=="CR"){
                                loanCR++;
                                }
                            else if(responselist[i].invType=="DI"){
                                invLOAN++;
                            }
                break;
                case "MISCBUY": if(responselist[i].invType=="CR"){
                                    miscbuyCR++;
                                }
                                else if(responselist[i].invType=="DI"){
                                    invMISCBUY++;
                                }
                break;
                case "SETTLEMENT": if(responselist[i].invType=="CR"){
                                        settlementCR++;
                                    }
                                    else if(responselist[i].invType=="DI"){
                                        invSETTLEMENT++;
                                    }
                break;
                default: break;
            }
            worksheet[i]={
                    "Invoice Number" : responselist[i].invNumber,
                    "Type" : invObj[0].id,
                    "Invoice/Credit/Debit": responselist[i].invType=="CR"?"Credit":responselist[i].invType=="DR"?"Debit":"Invoice",                    
                    "Creation Time (UTC)" : responselist[i].createdOn
                    };
        }
    }
    }
//#endregion

//#region GR
let grCount=0;
if(assetName=='GR'){    
    for(let i=0;i<responselist.length;i++){     //looping through fetched JSON result            
        let grcptType=responselist[i].ref.filter(e=>e.idQualf=='ZF' && e.id!='SI');
        if(!_undrscr.isEmpty(grcptType)){       //checking if GR response list is valid and not empty
            worksheet[grCount]={
                    "Good Receipt Number" : responselist[i].grNumber,
                    "Creation Time (UTC)" : responselist[i].createdOn
                    };
                grCount++;
            }
        }
    }
//#endregion

//#region Payment
if(assetName=='Payment'){
    for(let i=0;i<responselist.length;i++){   //looping through fetched JSON result            
        if(!_undrscr.isEmpty(responselist[i])){       //checking if Payment response list is valid and non-empty

            worksheet[i]={
                    "Payment Number" : responselist[i].paymentNumber,
                    "Creation Time (UTC)" : responselist[i].createdOn
                    };
            }
        }
    }
//#endregion

    //Save the workbook
     xlsx.utils.sheet_add_json(workbook.Sheets[assetName], worksheet)     
     
     switch(assetName){
         case 'PO': //arrCount[0]={"":"Total PO","Count":responselist.length};
        //            arrCount[1]={"":"Loan PO","Count":LOANpo};
        //            arrCount[2]={"":"T2 PO","Count":T2po};
        //            arrCount[8]={"  ":"SI PO","Count":SIpo};
                console.log("Total PO:"+responselist.length+"\n"+"Loan PO:"+LOANpo+"\nT2 PO:"+T2po+"\nSI PO:"+SIpo);
        break;
        case 'Shipment': //arrCount[9]={"":"Total Shipment","Count":responselist.length};
                        console.log("Total Shipment:"+responselist.length);
        break;
        case 'Invoice': //arrCount[4]={"":"Total Invoice","Count":responselist.length};
        //                 arrCount[5]={"":"T2 Invoice","Count":invT2};
        //                 arrCount[6]={"":"LOAN Invoice","Count":invLOAN};
        //                 arrCount[10]={"":"SI Invoice","Count":invSI};
        //                 arrCount[11]={"":"SETTLEMENT Invoice","Count":invSETTLEMENT};
        //                 arrCount[13]={"":"MISCBUY Invoice","Count":invMISCBUY};
        //                 arrCount[7]={"":"LOAN Credit","Count":loanCR};
        //                 arrCount[12]={"":"SETTLEMENT Credit","Count":settlementCR};
                        console.log("Total Invoice:"+responselist.length+"\nSETTLEMENT Invoice:"+invSETTLEMENT+"\nLOAN Invoice:"+invLOAN+"\nT2 Invoice:"+invT2+"\nSI Invoice:"+invSI+"\nMISCBUY Invoice:"+invMISCBUY);
                        console.log("LOAN Credit:"+loanCR+"\nSETTLEMENT Credit:"+settlementCR);
        break;
        case 'GR': //arrCount[3]={"":"Total GR","Count":grCount};
                    console.log("Total GR:"+grCount);
        break;
        case 'Payment': console.log("Total Payment:"+responselist.length);
        break;
        default: 
            break;
     }
     //xlsx.utils.sheet_add_json(workbook.Sheets["Total Count"], arrCount)  
     xlsx.writeFile(workbook, 'Weekly Asset Report.xlsx');
     
    }
    catch(e){
        console.log(`Error in ${assetName} module:`,e)
    }
}
    
