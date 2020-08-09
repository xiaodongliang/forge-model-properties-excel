const express = require('express'); 
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { AuthClientTwoLegged, DerivativesApi } = require('forge-apis');
const config = require('./config');

let forge_auth = new AuthClientTwoLegged(config.client_id, config.client_secret,config.scopes);
let forge_deriv = new DerivativesApi();
 

const {FORGE_CLIENT_ID, FORGE_CLIENT_SECRET, FORGE_MODEL_URN } = process.env;
if (!FORGE_CLIENT_ID || !FORGE_CLIENT_SECRET || !FORGE_MODEL_URN ) {
    console.error('Some of the following env. variables are missing:');
    console.error('FORGE_CLIENT_ID, FORGE_CLIENT_SECRET, MODEL_URN');
    return;
} 

(async()=>{ 
    try{
        const authRes = await forge_auth.authenticate(config.scopes);

        //this script assumes the model has been translated.
        //or comment out the lines to upload the local file to translate..
        
        const manifestRes = await forge_deriv.getManifest(config.model_urn,{},forge_auth,forge_auth.getCredentials());
        if(manifestRes.body.status == 'success' && manifestRes.body.progress == 'complete'){
            const metaRes = await forge_deriv.getMetadata(config.model_urn,{},forge_auth,forge_auth.getCredentials());
            const meta3DviewUrn = metaRes.body.data.metadata.find(i=>i.name == '{3D}' && i.role == '3d');
            if(meta3DviewUrn){
                let propertiesRes = {body:{result:'success'}}

                while(propertiesRes.body.result == 'success'){

                    propertiesRes.result == null
                    propertiesRes = await forge_deriv.getModelviewProperties(config.model_urn,meta3DviewUrn.guid,{forceget:true},forge_auth,forge_auth.getCredentials());
                    
                }

                const allProperties = propertiesRes.body.data.collection 

                const workbook = new ExcelJS.Workbook();
                var worksheet = workbook.addWorksheet('powerbi_forge')  
                
                // const sheetColumns = [
                //     { id: 'dbid',        columnTitle: 'DbId',              columnWidth: 8,     locked: true },
                //     //{ id: 'category',        columnTitle: 'Category',           columnWidth: 16,    locked: false },
                //     { id: 'material',    columnTitle: 'Material',     columnWidth: 16,    locked: false }
                    
                // ]; 
                //worktable.columns = sheetColumns.map(col => {
                //    return { key: col.id, header: col.columnTitle, width: col.columnWidth };
                //});

                const columns =  [
                    {name: 'dbid', totalsRowLabel: 'dbid:', filterButton: true},
                    {name: 'material', totalsRowFunction: 'material', filterButton: false},
                  ]

                let rows = [];
                allProperties.forEach(async p =>  {
                    //let hasMaterial = p.properties.find(i=>i.attributeName && i.attributeName.includes('Material'))
                    //let hasCategory= p.properties.find(i=>i.attributeName && i.attributeName.includes('Category'))

                     let hasMaterial = p.properties['Materials and Finishes']
                    //let hasCategory= p.properties['Materials and Finishes']

                    if(hasMaterial ){//&& hasCategory){
                        let eachrow = []
                        rows['dbid'] = p.objectid;
                        //rows['category'] = hasCategory.displayValue; 
                        rows['material'] = hasMaterial['Material'];
                        rows.push(eachrow);  
                    } 
                }); 
 
                var worktable = worksheet.addTable({
                    name: 'powerbi_forge',
                    ref: 'A1',
                    headerRow: true,
                    totalsRow: true,
                    columns:columns,
                    rows:rows
                  }
                )

                const buffer = await workbook.xlsx.writeBuffer();
                fs.writeFile("powerbi_forge.xlsx", buffer, "binary",function(err) {
                    if(err) {
                    console.log(err);
                    } else {
                    console.log("The Excel file was saved!");
                    }
                })  

            }else{
                console.error('this model has no {3D} view') 
            }
        }else{
            console.error('this model manifest is not available..')
        }

    }
    catch(e){
        console.error(e);
    }  
})()
 