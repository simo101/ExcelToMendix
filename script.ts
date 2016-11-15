/// <reference path='./typings/tsd.d.ts' />

import { MendixSdkClient, OnlineWorkingCopy, Project, Revision, Branch, loadAsPromise } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus } from "mendixmodelsdk";
import when = require('when');
import XLSX = require('xlsx');
const jsonObj = {};

const username = "simon.black@mendix.com";
const apikey = "436a5070-72c7-458e-a372-1ecc3598cb7d";
const projectId = "fd4be9ba-23f0-40ba-8ad9-6c9372e73f74";
const projectName = "ConvertFromExcel";
const moduleName = "Navigation";
const revNo = -1; // -1 for latest
const branchName = null // null for mainline
const wc = null;
const client = new MendixSdkClient(username, apikey);


/*
 * PROJECT TO ANALYZE
 */
const project = new Project(client, projectId, projectName);

client.platform().createOnlineWorkingCopy(project, new Revision(revNo, new Branch(project, branchName)))
    .then(OnlineWorkingCopy =>
    pickExcelDocument())
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
function pickDomainModel(workingCopy:OnlineWorkingCopy): domainmodels.IDomainModel{
    return workingCopy.model().allDomainModels().filter(domainModel =>{
        return domainModel.moduleName === moduleName;
    })[0];
}

function pickExcelDocument(): when.Promise<void> {
    var workbook = XLSX.readFile('test.xlsx');
    var sheet_name_list = workbook.SheetNames;
    sheet_name_list.forEach(function (y) { /* iterate through sheets */
        var worksheet = workbook.Sheets[y];
        var z;
        for (z in worksheet) {
            /* all keys that do not begin with "!" correspond to cell addresses */
            if (z[0] === '!') continue;
            console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
        }
    });


    return;
}
