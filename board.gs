var SS=SpreadsheetApp.getActiveSpreadsheet();
var PROJECT_SHEET_NAME='PRIORITY - ALL'
var trello_board_id="5ecd07936438844a8863105a";
var trello_key="fad1c87c9cc4415a9162425d93432f44";
var trello_token="09fd3c4307681ed7c099233c40a7c3ca3d1ba7192b5b9cac6fba6451449236a7";
function createConfiguration()
{
  var url='https://api.trello.com/1/boards/'+trello_board_id+'/lists/open?';
  url+="key=" + trello_key + "&token=" + trello_token;
//  https://api.trello.com/1/boards/5ecd07936438844a8863105a/lists/open?key=fad1c87c9cc4415a9162425d93432f44&token=09fd3c4307681ed7c099233c40a7c3ca3d1ba7192b5b9cac6fba6451449236a7
  var response=UrlFetchApp.fetch(url).getContentText();
  
  var json=JSON.parse(response);
  var projectListId='',inProgressListId='';
  for(var i in json)
  {
    if(json[i].name=="Projects")
    {      
      projectListId=json[i].id;      
    }
    else if(json[i].name=="In Progress")
      inProgressListId=json[i].id;
  }
  var folderId=getFieldsFromBoard();
  
  var output=[];
  output.push([projectListId]);
  output.push([inProgressListId]);
  output.push([folderId]);
  output.push([getPlugin()]);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange(1,2,output.length).setValues(output);
  
  
}

// Priority Tasks
function getCardsByList()
{
  var configureData=SS.getSheetByName("Configuration").getDataRange().getValues();
  var projectListId=configureData[0][1];
  var folderId=configureData[2][1];
  
  var url='https://trello.com/1/lists/'+projectListId+'/cards?fields=id,name,due&customFieldItems=true'
  url+="&key=" + trello_key + "&token=" + trello_token;
  
  var response=UrlFetchApp.fetch(url).getContentText();  
  var json=JSON.parse(response);
  
  var today=new Date();
  var outputData=[];
  var dueDates=[],projectNames=[],folderNumbers=[];
  var shProjects=SS.getSheetByName(PROJECT_SHEET_NAME);
  var lastRow=shProjects.getLastRow();
  var projectData=shProjects.getRange(1,6,lastRow).getValues();
  var projects=[];
  var startRowNumber=5;
  for(var i in projectData)
  {
    if(projectData[i][0]=="")
      break;
    projects.push(projectData[i][0]);    
  }
  lastRow=Number(i);  

//  var folderId='5dd004b9e1f7ed4d7a26372d';
  
  for(var i in json)
  {
    var name=json[i].name;  
    var customFields=json[i].customFieldItems;
    var folderNumber='';
    var date=new Date(json[i].due);
    for(var j in customFields)
    {
      if(customFields[j].idCustomField==folderId)
      {
        folderNumber=customFields[j].value.number;
        break;
      }
    }
    
    var index=projects.indexOf(json[i].name);
    if(index!=-1)
    {
      if(today.getTime()>date.getTime())
      {
        shProjects.getRange(index+1,6).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID_MEDIUM);      
      }
      else
      {
        shProjects.getRange(index+1,6).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID);
      }
      shProjects.getRange(index+1,3).setValue(date);
      shProjects.getRange(index+1,8).setValue(folderNumber);
      continue;
    }
    
    dueDates.push([date]);
    projectNames.push([json[i].name]);
    folderNumbers.push([folderNumber]);
  }
  
  if(projectNames.length>0)
  {
    lastRow+=1;
    shProjects.getRange(lastRow,3,dueDates.length).setValues(dueDates);
    shProjects.getRange(lastRow,6,dueDates.length).setValues(projectNames);
    shProjects.getRange(lastRow,8,dueDates.length).setValues(folderNumbers);
  }
  SpreadsheetApp.flush();
  unformatProjectCards()
}



function getFieldsFromBoard( ) {
  var url = "https://api.trello.com/1/boards/"+trello_board_id+"/customFields?" + 
    "key=" + trello_key + "&token=" + trello_token;
  
  try {
    var response = UrlFetchApp.fetch( url );
  } catch (err) {
    console.error( err );
    return {};
  }
  
  var json = JSON.parse(response.getContentText());
  for(var i in json)
  {
    if(json[i].name=="Folder #")    
      return json[i].id;    
  }
}

function getPlugin()
{
  var url='https://api.trello.com/1/boards/'+trello_board_id+'/plugins?filter=enabled&fields=id,name';
  url+="&key=" + trello_key + "&token=" + trello_token;
  var response=UrlFetchApp.fetch(url).getContentText();
  
  var json=JSON.parse(response);
  var pluginId='';
  
  for(var i in json)
  {
    if(json[i].name=="Time in List")
      return json[i].id;
  }
  return pluginId;  
}

function test()
{
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  for(var i=1;i<200;i++)
    sheet.getRange(i,3).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID);
  
}

function getBoardLists()
{
  var url='https://api.trello.com/1/boards/'+trello_board_id+'/lists/open';
  url+="?key=" + trello_key + "&token=" + trello_token;

  var response=UrlFetchApp.fetch(url).getContentText();
  
  var json=JSON.parse(response);
  
  var listData=[];
  
  for(var i in json)
  {
    listData.push([json[i].id,json[i].name]) ;
  }
  if(listData.length>0)
    SS.getSheetByName("Testing Sheet").getRange(1,8,listData.length,2).setValues(listData);
  


}

   

function unformatProjectCards()
{
//  return;
  var configureData=SS.getSheetByName("Configuration").getDataRange().getValues();
  var projectListId=configureData[0][1];
  var folderId=configureData[2][1];
  
  var url='https://trello.com/1/lists/'+projectListId+'/cards?fields=id,name,due&customFieldItems=true'
  url+="&key=" + trello_key + "&token=" + trello_token;
  
  var response=UrlFetchApp.fetch(url).getContentText();  
  var json=JSON.parse(response);
  
  var currentProjects=[];
  
  for(var i in json)
    currentProjects.push(json[i].name);
  
  var shPriority=SS.getSheetByName("PRIORITY - ALL");  
  var existingProjects=shPriority.getRange(1,6,shPriority.getLastRow()).getValues();
  
//  var outputArr=[];
  
  
  for(var i=4;i<existingProjects.length;i++)
  {
    var index=currentProjects.indexOf(existingProjects[i][0]);
    if(index==-1)
      shPriority.getRange(i+1,6).setBorder(true, true, true, true, true, true, null, null);
    else      
      shPriority.getRange(i+1,6).setBorder(true,true,true,true,true,true,"black",SpreadsheetApp.BorderStyle.SOLID_THICK);    
//    return;    
  }
//  var len1=currentProjects.length;
//  var len2=existingProjects.length;
//  var len3=outputArr.length;
//  return;
  
  
}