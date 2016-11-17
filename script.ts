/// <reference path='./typings/tsd.d.ts' />

import { MendixSdkClient, OnlineWorkingCopy, Project, Revision, Branch, loadAsPromise } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus } from "mendixmodelsdk";
import { MendixModelComponents} from "mendixmodelcomponents";
import when = require('when');
import XLSX = require('xlsx');
import {IWorkSheet, IWorkBook} from "xlsx";
const jsonObj = {};

const username = "simon.black@mendix.com";
const apikey = "436a5070-72c7-458e-a372-1ecc3598cb7d";
const projectId = "fd4be9ba-23f0-40ba-8ad9-6c9372e73f74";
const projectName = "ConvertFromExcel";
const moduleName = "MyFirstModule";
const revNo = -1; // -1 for latest
const branchName = null // null for mainline
const wc = null;
const client = new MendixSdkClient(username, apikey);
var components = null;
/*
 * PROJECT TO ANALYZE
 */
const project = new Project(client, projectId, projectName);

client.platform().createOnlineWorkingCopy(project, new Revision(revNo, new Branch(project, branchName)))
    .then(workingCopy => {
        components = new MendixModelComponents(workingCopy.model());
        return pickDomainModel(workingCopy).then(domainModel => {
            return createEntities(components, domainModel);
        }).then(_=>{
            return workingCopy;
        });
    }).then(workingCopy => {
        workingCopy.commit();
    })
    .done(

    () => {
        console.log("Done.");
    },
    error => {
        console.log("Something went wrong:");
        console.dir(error);
    }
    );

/*
* This function picks the right domain model based on the module name provided.
*/
function pickDomainModel(workingCopy: OnlineWorkingCopy): when.Promise<domainmodels.DomainModel> {
    return loadAsPromise(workingCopy.model().allDomainModels().filter(domainModel => {
        return domainModel.moduleName === moduleName;
    })[0]);
}

function createEntities(components: MendixModelComponents, domainModel: domainmodels.DomainModel): when.Promise<void> {
    var workbook = XLSX.readFile('test.xlsx');
    var sheet_name_list = workbook.SheetNames;
    var firstSheet = sheet_name_list[0];
    // sheet_name_list.forEach(function (y) { /* iterate through sheets */
    var worksheet = workbook.Sheets[firstSheet];
    var z;
    var x, y = 100;
    for (z in worksheet) {
        /* all keys that do not begin with "!" correspond to cell addresses and first header row */
        if (z[0] === '!' || z === 'A1' || z === 'B1' || z === 'C1' || z === 'D1' || z === 'E1') continue;
        // console.log(firstSheet + "!" + z + "=" + JSON.stringify(worksheet[z].v));
        if (z.startsWith('A')) {
            components.createEntity(domainModel, worksheet[z].v, x, y);
            x += 50;
            y += 50;
        }

    }
    // });


    return;
}
