var runflag=false;
var breakflag=false;
var api_token;
function onInstall(e) {
    onOpen(e);
}

// trello variables
var board_id = "5ecd07936438844a8863105a";
var url = "https://api.trello.com/1/";

//  https://api.trello.com/1/boards/5ecd07936438844a8863105a/lists/all/?key=fad1c87c9cc4415a9162425d93432f44&token=09fd3c4307681ed7c099233c40a7c3ca3d1ba7192b5b9cac6fba6451449236a7

var run_script=false;

function doGet() {
    return HtmlService.createHtmlOutputFromFile('page');
}

function update_board_id(b_id){
    board_id=b_id;
}
function getUnreadEmails() {
    return"my return";
}

function getTrelloBoard(license_key) {
//    var response;
//    Logger.log(license_key);
//
//    response = UrlFetchApp.fetch('http://datahighlighter.com/api/getTrelloBoard/' + license_key)
//    
//    if(response != []) {
//        SpreadsheetApp.getActiveSheet().getRange(41, 1).setValue(response);
//        document.getElementById("connect_board").removeAttribute("disabled");
//    } else {
//        SpreadsheetApp.getActiveSheet().getRange(42, 1).setValue("faild");
//    }
    
}

function connect_trello(){
    return "asdf";
}
var tiger=true;
function createTimeDrivenTriggers() {
SpreadsheetApp.getActiveSheet().getRange(1, 1).setValue("111");
    var triggers=ScriptApp.getProjectTriggers();
    for(var i in triggers)
        ScriptApp.deleteTrigger(triggers[i]);

    // Trigger every 10 second
    ScriptApp.newTrigger('Time_call')
        .timeBased()
        .everyMinutes(1)
        .create();
}

function Sync(){
    if(tiger) {
        tiger=false;
        createTimeDrivenTriggers();
    }
    else{
        tiger= true;
        deleteTrigger();
    }
}

function main(b_id, key_and_token) {
    //getMembers();
    var listnames=[];
    labels_All=[];
    var ss = SpreadsheetApp.getActiveSheet();
    ss.clear();
    ss.getRange('A2:A').activate();
    
    ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('A1:A1').activate();
    ss.appendRow(["Sync'D", "CARD NAME","List Name", "Due Date", "Labels", "Last Time", "Last Person","Text Content"]);
    var test_response = 1;
    var response = UrlFetchApp.fetch(url + "boards/" + b_id + "/lists/all/?" + key_and_token);
  
    console.log( "response: " + response );
    var lists = JSON.parse((response.getContentText()));
    var p=0;
      
    for (var i = 0; i < lists.length; i++) {
        var list = lists[i];
        var response = UrlFetchApp.fetch(url + "list/" + list.id + "/cards/all?" + key_and_token);
        var cards = JSON.parse(response.getContentText());
        if (!cards)
            continue;
        listnames.push(list.name);
        for (var j = 0; j < cards.length; j++) {
            var card = cards[j];
            cardIDs.push(card.id);
            Logger.log(cards);
            var response = UrlFetchApp.fetch(url + "cards/" + card.id + "/actions/?" + key_and_token);
            var carddetails = JSON.parse(response.getContentText());
    
            Logger.log(carddetails);
            var name = card.name;
            var link = card.url;
            var listname = list.name;
            var archivedList = card.closed;

            var txt="";
            for (var k = 0; k < carddetails.length; k++) {        
                var dato = carddetails[0].date;
                var fullname = carddetails[0].memberCreator.fullName;

                if(carddetails[k].data.text!=null)
                    txt+=carddetails[k].data["text"]+"\n";
            }
            var lbl="";
            var lblflag;
            for (var pp=0;pp<card["labels"].length;pp++){
                lbl+=card["labels"][pp]["name"]+'\n';
                var lbltmp=card["labels"][pp]["name"];
                //ss.getRange(1,1).setValue(lbl)
                lblflag=true;

                for(var t=0;t<labels_All.length;t++){                  
                   if(archivedList){
                        ss.getRange(p+2,3,1,6).setBackground("#999999");
                    }
                    if(labels_All[t]==lbltmp){
                        lblflag=false;
                        break;
                    }
                }
                if(lblflag){
                    labels_All.push(lbltmp);
                    //ss.getRange(1, 1).setValue(lbltmp);
                }
            }
            ss.appendRow(['', name,listname, card["due"], lbl, dato,fullname,txt]);
            ss.getRange(p+2, 1,1,1).insertCheckboxes();
            ss.getRange(p+2, 1,1,1).setValue(true);
            p+=1
        }                                       
    }
    var retval=[listnames,labels_All];
    return retval;
}

function condition_1(lab, val,ind, con) {
    var ss = SpreadsheetApp.getActiveSheet();

  
    try{
        for (var i = 0; i < con.length; i++) {
            var arr = con[i];
        //ss.getRange(1,1).setValue(val);
        
            if (arr.fieldoption[0] == lab) {
            // ss.getRange(1,1).setValue(x.toString()+y.toString());
            
                var flag = false;
                if (arr.operatoroption[0] == "0") {
                    
                    if (val[ind] == arr.valuetxt[0]) {
                        flag = true;
                    
                    }
                
                
                }
                if (arr.operatoroption[0] == "1") {
                    if (val[ind] > arr.valuetxt[0]) {
                        flag = true;
                    }
                }
                if (arr.operatoroption[0] == "2") {
                    if (val[ind] < arr.valuetxt[0]) {
                        flag = true;
                    }
                }
                if (arr.operatoroption[0] == "3") {
                    if (val[ind] >= arr.valuetxt[0]) {
                        flag = true;
                    }
                }
                if (arr.operatoroption[0] == "4") {
                    if (val[ind] <= arr.valuetxt[0]) {
                        flag = true;
                    }
                }

                if(lab!="Due Date"){
                    if(val[ind].search(arr.valuetxt[0])>-1){
                    flag=true;
                    }
                }
                if (flag) {
                
                ///////////////////////////////
                    var val1,ind1;
                    
                    if(arr.fieldoption[1]=="Due Date")
                        ind1=0;
                    if(arr.fieldoption[1]=="Label")
                        ind1=1;
                    if(arr.fieldoption[1]=="Last Person Connected")
                        ind1=3;
                    if(arr.fieldoption[1]=="Assignee")
                        ind1=4;
                    if(arr.fieldoption[1]=="Text Content")
                        ind1=5;
                    if(arr.fieldoption[1]=="Last Time")
                        ind1=2;
                    val1=val[ind1];// ss.getRange(1,1).setValue(x.toString()+y.toString()+flag+arr.andor+val[2]);
                    ////////////////////////
                    // ss.getRange(1,1).setValue(arr.valuetxt[1]);
                    if (arr.andor == "And") {
                        flag = false;
                    
                        if (arr.operatoroption[1] == "0") {
                            if (val1 == arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "1") {
                            if (val1 > con.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "2") {
                            if (val1 < arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "3") {
                            if (val1 >= con.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "4") {
                            if (val1 <= arr.valuetxt[1]) {
                                flag = true;
                            }
                        } 
                        if(arr.fieldoption[1]!="Due Date"  && arr.operatoroption[1] !="0"){
                        // ss.getRange(1,1).setValue(val1);
                            if(val1.search(arr.valuetxt[1])>-1){
                            flag=true;
                        //     ss.getRange(1,1).setValue(x.toString()+y.toString());
                            }
                        }
                    }
                } else {
                    if (arr.andor == "Or") {
                    ///////////////////////////////
                        var val1,ind1;
                        
                        if(arr.fieldoption[1]=="Due Date")
                            ind1=0;
                        if(arr.fieldoption[1]=="Label")
                            ind1=1;
                        if(arr.fieldoption[1]=="Last Person Connected")
                            ind1=3;
                        if(arr.fieldoption[1]=="Assignee")
                            ind1=4;
                        if(arr.fieldoption[1]=="Text Content")
                            ind1=5;
                        if(arr.fieldoption[1]=="Last Time")
                            ind1=2;
                        val1=val[ind1];// ss.getRange(1,1).setValue(x.toString()+y.toString()+flag+arr.andor+val[2]);
                        ////////////////////////
                        flag = false;
                        if (arr.operatoroption[1] == "0") {
                            if (val1 == arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "1") {
                            if (val1 > arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "2") {
                            if (val1 < arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "3") {
                            if (val1 >= arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "4") {
                            if (val1 <= arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if(arr.fieldoption[1]!="Due Date"){
                            if(val1.search(arr.valuetxt[1])>-1){
                            flag=true
                            }
                        }
                    }
                }
                return flag;
    //          if(flag){
    //            
    //            ss.getRange(x,y).setBackground(arr.backcolor);
    //            ss.getRange(x,y).setBorder(true, null, true, null, false, false, arr.bordercolor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    //            ss.getRange(x,y).setFontColor(arr.textcolor);
    //            return;
    //          }
            }
        }
    } catch(e) {
        Logger.log(e);
    }
  return false;
}

function sync_main(searchtxt,conditionarray_1, b_id, key_and_token) {

    var arr=searchtxt.split(",");
    if(arr.length==0) return;
    var ss = SpreadsheetApp.getActiveSheet();
    //    ss.getRange(1, 1).setValue("123");
    
    var startrow=2;
    var lastrow=ss.getLastRow();
    var data = ss.getRange(startrow, 1, lastrow-startrow+1,10).getValues();
    var data_allow=[];
    for(var i=0;i<data.length;i++){
        if(data[i][0])
            data_allow.push(data[i]);
    }

    ss.clear();
    ss.getRange('A2:A').activate();
    ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('A1:A1').activate();
    ss.appendRow(["Sync'D", "Card Nmae","ListName", "Due Date", "Labels", "Last Time", "Last Person","Text Content"]);
    var response = UrlFetchApp.fetch(url + "boards/" + b_id + "/lists/all/?" + key_and_token);
    var lists = JSON.parse((response.getContentText()));
    var p=0;
    for (var i = 0; i < lists.length; i++) { // ss.getRange(1, 1).setValue("123");
        var list = lists[i];
        var flag=false;
        for(var o=0;o<arr.length-1;o++){
            if(list.name.search(arr[o])>-1){
                flag=true;
                break;
            }
        }
        if(flag){
            var response = UrlFetchApp.fetch(url + "list/" + list.id + "/cards/all?" + key_and_token);
            var cards = JSON.parse(response.getContentText());
            if (!cards)
            continue;

            for (var j = 0; j < cards.length; j++) {
                var card = cards[j];
                cardIDs.push(card.id);
                Logger.log(cards);
                var response = UrlFetchApp.fetch(url + "cards/" + card.id + "/actions/?" + key_and_token);
                var carddetails = JSON.parse(response.getContentText());
            //      if (!carddetails)
            //        continue;
            
                Logger.log(carddetails);
                var name = card.name;
                var link = card.url;
                var listname = list.name;
                var txt="";
                for (var k = 0; k < carddetails.length; k++) {        
                    var dato = carddetails[0].date;
                    var fullname = carddetails[0].memberCreator.fullName+'\n';

                    if(carddetails[k].data.text!=null)
                    txt+=carddetails[k].data["text"]+"\n";
                }

                var lbl="";
                for (var pp=0;pp<card["labels"].length;pp++)
                    lbl+=card["labels"][pp]["name"]+'\n';

                var due="";
                if(card["due"]!=null && card["due"].length>10)
                    due=card["due"].substr(0,10);

                var valarr=[ due,lbl,dato,fullname,txt];
                
                if(conditionarray_1.length>0){
                    if(condition_1("Due Date", valarr,0, conditionarray_1) || condition_1("Label", valarr,1, conditionarray_1)){
                    
                    
                        ss.appendRow(['', name,listname, card["due"], lbl, dato,fullname,txt]);
                        
                        ss.getRange(p+2, 1,1,1).insertCheckboxes();
                        
                        ss.getRange(p+2, 1,1,1).activate();
                        ss.getRange(p+2, 1,1,1).setDataValidation(SpreadsheetApp.newDataValidation()
                        .setAllowInvalid(false)
                        .setHelpText('Enter Checked or Unchecked.')
                        .requireCheckbox('Checked', 'Unchecked')
                        .build());
                        
                        for(var k=0;k<data.length;k++){
                            if(data[k][0] && data[k][1]==name){
                                ss.getRange(p+2, 1,1,1).setValue(true);
                            }
                        }

                        if(data.length==0)
                            ss.getRange(p+2, 1,1,1).setValue(true);
                        p+=1;
                    }
                }else{
                    ss.appendRow(['', name,listname, card["due"], lbl, dato,fullname,txt]);
                    
                    ss.getRange(p+2, 1,1,1).insertCheckboxes();
                    for(var k=0;k<data.length;k++){
                        if(data[k][0] && data[k][1]==name){
                            ss.getRange(p+2, 1,1,1).setValue(true);
                        }
                    }
                    if(data.length==0)
                        ss.getRange(p+2, 1,1,1).setValue(true);
                    
                    p+=1;
                }

            }
        
        }                                       
    }
}

function pull(searchtxt, key_and_token) {
  
    var arr=searchtxt.split(",");
    var ss = SpreadsheetApp.getActiveSheet();
    // ss.getRange(1,1).setValue(searchtxt);
    var startrow=2;
    var lastrow=ss.getLastRow();
    var data = ss.getRange(startrow, 1, lastrow-startrow+1,10).getValues();
    var data_allow=[];
    for(var i=0;i<data.length;i++){
        if(data[i][0])
            data_allow.push(data[i]);
    }
  
    var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/lists/all/?" + key_and_token);
    var lists = JSON.parse((response.getContentText()));

    var p=0;
    
    for (var i = 0; i < lists.length; i++) {
        var list = lists[i];
        var flag=false;
        var response = UrlFetchApp.fetch(url + "list/" + list.id + "/cards/all?" + key_and_token);
        var cards = JSON.parse(response.getContentText());

        if (!cards)
            continue;

        for (var j = 0; j < cards.length; j++) {
            var card = cards[j];

            var response = UrlFetchApp.fetch(url + "cards/" + card.id + "/actions/?" + key_and_token);
            var carddetails = JSON.parse(response.getContentText());
            // if (!carddetails)
            //   continue;
            Logger.log(carddetails);
            var name = card.name;
            var link = card.url;
            var listname = list.name;
            var txt="";
            for (var k = 0; k < carddetails.length; k++) {        
                var dato = carddetails[0].date;
                var fullname = carddetails[0].memberCreator.fullName;

                if(carddetails[k].data.text!=null)
                    txt+=carddetails[k].data["text"]+"\n";
            }
            var lbl="";
            for (var pp=0;pp<card["labels"].length;pp++)
                lbl+=card["labels"][pp]["name"]+"\n";

            var flagvar=true;
            var flag=true;
            for(var o=0;o<arr.length;o++){
                
                if(arr[o]==listname){
                    flag=false;
                    break;
                }
            }

            for(var k=0;k<data.length;k++){
                
                if(data[k][1]==name){
                    flagvar=false;break;
                }
            }

            if(flagvar && flag==false){
                ss.appendRow(['', name,listname, card["due"], lbl, dato,fullname,txt]);
                ss.getRange(ss.getLastRow(), 1,1,1).insertCheckboxes();
                ss.getRange(ss.getLastRow(), 1,1,1).setValue(true);

            }

        }
    }
}

function Time_call(searchtxt, key_and_token){
    if(runflag==true){
        breakflag=true;
        while(runflag){
      
        }
    }
    runflag=true;
    var arr=searchtxt.split(",");
    var ss = SpreadsheetApp.getActiveSheet();
    var startrow=2;
    var lastrow=ss.getLastRow();
    var data = ss.getRange(startrow, 1, lastrow-startrow+1,10).getValues();
    var data_allow=[];
    for(var i=0;i<data.length;i++){
        if(data[i][0])
            data_allow.push(data[i]);
    }

  
    var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/lists/all/?" + key_and_token);
    var lists = JSON.parse((response.getContentText()));
        
    var p=0;
      
    for (var i = 0; i < lists.length; i++) {
        var list = lists[i];
        var flag=false;

        var response = UrlFetchApp.fetch(url + "list/" + list.id + "/cards/all?" + key_and_token);
        var cards = JSON.parse(response.getContentText());
        if (!cards)
            continue;

        for (var j = 0; j < cards.length; j++) {
            var card = cards[j];
            cardIDs.push(card.id);
            Logger.log(cards);
            var response = UrlFetchApp.fetch(url + "cards/" + card.id + "/actions/?" + key_and_token);
            var carddetails = JSON.parse(response.getContentText());

            Logger.log(carddetails);
            var name = card.name;
            var link = card.url;
            var listname = list.name;
            var archivedList = card.closed;

            var txt="";
            for (var k = 0; k < carddetails.length; k++) {        
                var dato = carddetails[0].date;
                var fullname = carddetails[0].memberCreator.fullName;

            if(carddetails[k].data.text!=null)
                txt+=carddetails[k].data["text"]+"\n";
            }
      
            var lbl="";
            for (var pp=0;pp<card["labels"].length;pp++)
                lbl+=card["labels"][pp]["name"]+"\n";

            //ss.appendRow(['', name,listname, card["due"], lbl, dato,fullname,txt]);
            for(var k=0;k<data.length;k++){
                if(archivedList){
                    ss.getRange(p+2,3,1,6).setBackground("#999999");
                }
                if(data[k][0] && data[k][1]==name){
                    if(listname!=ss.getRange(k+startrow, 3).getValue())
                        ss.getRange(k+startrow, 3).setValue(listname);
                    if(card["due"]!=ss.getRange(k+startrow, 4).getValue())
                        ss.getRange(k+startrow, 4).setValue(card["due"]);
                    if(lbl!=ss.getRange(k+startrow, 5).getValue())
                        ss.getRange(k+startrow, 5).setValue(lbl);
                    if(dato!=ss.getRange(k+startrow, 6).getValue())
                        ss.getRange(k+startrow, 6).setValue(dato);
                    if(fullname!=ss.getRange(k+startrow, 7).getValue())
                        ss.getRange(k+startrow, 7).setValue(fullname);
                    if(txt!=ss.getRange(k+startrow, 8).getValue())
                        ss.getRange(k+startrow, 8).setValue(txt);
                }
            }
        p+=1
        }
    }
    runflag=false;
}

function apply(conditionarray, key_and_token){
    if(runflag==true){
        breakflag=true;
        while(runflag){

        }
    }
  
    runflag=true;
    members_all=getMembers(key_and_token);
    var ss = SpreadsheetApp.getActiveSheet();
  
    var startrow=2;
    var lastrow=ss.getLastRow();
    var data = ss.getRange(startrow, 1, lastrow-startrow+1,10).getValues();
    var data_allow=[];

    for(var i=0;i<data.length;i++){
        if(data[i][0])
            data_allow.push(data[i]);
    }
  
    var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/lists/all/?" + key_and_token);
    var lists = JSON.parse((response.getContentText()));
    
    var p=0;
      
    for (var i = 0; i < lists.length; i++) {
        var list = lists[i];
        var flag=false;
        var response = UrlFetchApp.fetch(url + "list/" + list.id + "/cards/all?" + key_and_token);
        var cards = JSON.parse(response.getContentText());
        if (!cards)
            continue;

        for (var j = 0; j < cards.length; j++) {
            var card = cards[j];
        
            Logger.log(cards);
            var response = UrlFetchApp.fetch(url + "cards/" + card.id + "/actions/?" + key_and_token);
            var carddetails = JSON.parse(response.getContentText());
            console.log(carddetails);

            Logger.log(carddetails);
            var name = card.name;
            var link = card.url;
            var listname = list.name;
            var archivedList = card.closed;
            var joinname=''; 
            
            for (var oo=0;oo<card["idMembers"].length;oo++){
            
                for(var pp=0;pp<members_all.length;pp++){
                    if(card["idMembers"][oo]==members_all[pp]["id"]){
                            joinname+=members_all[pp]["fullName"]+"\n";
                            break;
                    }
                }
            }

            var txt="";
            for (var k = 0; k < carddetails.length; k++) {        
                var dato = carddetails[0].date;
                var fullname = carddetails[0].memberCreator.fullName;

                if(carddetails[k].data.text!=null)
                    txt+=carddetails[k].data["text"]+"\n";
            }

            var lbl="";
            for (var pp=0;pp<card["labels"].length;pp++)
                lbl+=card["labels"][pp]["name"]+"\n";
            //ss.appendRow(['', name,listname, card["due"], lbl, dato,fullname,txt]);
            for(var k=0;k<data.length;k++){
                if(archivedList){
                    ss.getRange(p+2,3,1,6).setBackground("#999999");
                }
                if(data[k][0] && data[k][1]==name){
                    if(listname!=ss.getRange(k+startrow, 3).getValue())
                        ss.getRange(k+startrow, 3).setValue(listname);
                    if(card["due"]!=ss.getRange(k+startrow, 4).getValue())
                        ss.getRange(k+startrow, 4).setValue(card["due"]);
                    if(lbl!=ss.getRange(k+startrow, 5).getValue())
                        ss.getRange(k+startrow, 5).setValue(lbl);
                    if(dato!=ss.getRange(k+startrow, 6).getValue())
                        ss.getRange(k+startrow, 6).setValue(dato);
                    if(fullname!=ss.getRange(k+startrow, 7).getValue())
                        ss.getRange(k+startrow, 7).setValue(fullname);
                    if(txt!=ss.getRange(k+startrow, 8).getValue())
                        ss.getRange(k+startrow, 8).setValue(txt);
                    if(joinname!=ss.getRange(k+startrow, 9).getValue())
                        ss.getRange(k+startrow, 9).setValue(joinname);
                    var due="";
                    if(card["due"]!=null && card["due"].length>10)
                        due=card["due"].substr(0,10);
                    var dato1=dato.substr(0,10);
                    var valarr=[ due,lbl,dato1,fullname,txt,joinname,listname];
                
                    if(due!=""){
                        condition("Due Date", valarr,0, conditionarray, k+startrow, 4);
                        //ss.getRange(1,4).setValue( card["due"].substr(0,10));
                    }
                    condition("Label", valarr,1, conditionarray, k+startrow, 5);
                    
                    condition("Last Time", valarr,2, conditionarray, k+startrow, 6);
                    
                    condition("Last Person Connected", valarr,3, conditionarray, k+startrow, 7);
                    //
                    condition("Text Content", valarr,4, conditionarray, k+startrow, 8);
                    condition("List Name", valarr,6, conditionarray, k+startrow, 3);
                    condition("Assignee", valarr,5, conditionarray, k+startrow, 9);
                    //ss.getRange(1,8).setValue( txt);
                }
            }
            p+=1
        }
    }
    runflag=false;
}

function condition(lab, val,ind, con, x, y) {
    var ss = SpreadsheetApp.getActiveSheet();
            ss.getRange(x,y).setBackground("#337abe");
            ss.getRange(x,y).setBorder(false, false, false, false, false, false, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            ss.getRange(x,y).setFontColor("#ffdf31");
  
    try{
        for (var i = 0; i < con.length; i++) {
            var arr = con[i];
        //ss.getRange(1,1).setValue(val);
        
            if (arr.fieldoption[0] == lab) {

                var flag = false;
                if (arr.operatoroption[0] == "0") {
                    if (val[ind] == arr.valuetxt[0]) {
                        flag = true;                    
                    }
                }
                if (arr.operatoroption[0] == "1") {
                    if (val[ind] > arr.valuetxt[0]) {
                        flag = true;
                    }
                }
                if (arr.operatoroption[0] == "2") {
                    if (val[ind] < arr.valuetxt[0]) {
                        flag = true;
                    }
                }
                if (arr.operatoroption[0] == "3") {
                    if (val[ind] >= arr.valuetxt[0]) {
                        flag = true;
                    }
                }
                if (arr.operatoroption[0] == "4") {
                    if (val[ind] <= arr.valuetxt[0]) {
                        flag = true;
                    }
                }

                if(lab!="Due Date"){
                    if(val[ind].search(arr.valuetxt[0])>-1){
                    flag=true;
                    }
                }

                if (flag) {
                    var val1,ind1;
                    
                    if(arr.fieldoption[1]=="Due Date")
                        ind1=0;
                    if(arr.fieldoption[1]=="Label")
                        ind1=1;
                    if(arr.fieldoption[1]=="Last Person Connected")
                        ind1=3;
                    if(arr.fieldoption[1]=="Assignee")
                        ind1=5;
                    if(arr.fieldoption[1]=="Text Content")
                        ind1=4;
                    if(arr.fieldoption[1]=="Last Time")
                        ind1=2;
                    if(arr.fieldoption[1]=="List Name")
                        ind1=6;
                    val1=val[ind1];
                    // ss.getRange(1,1).setValue(x.toString()+y.toString()+flag+arr.andor+val[2]);
                    ////////////////////////
                    // ss.getRange(1,1).setValue(arr.valuetxt[1]);
                    if (arr.andor[0] == "And") {
                        flag = false;
                    
                        if (arr.operatoroption[1] == "0") {
                            if (val1 == arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "1") {
                            if (val1 > con.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "2") {
                            if (val1 < arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "3") {
                            if (val1 >= con.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "4") {
                            if (val1 <= arr.valuetxt[1]) {
                                flag = true;
                            }
                        } 
                        if(arr.fieldoption[1]!="Due Date"){//  && arr.operatoroption[1] !="0"
                            // ss.getRange(1,1).setValue(val1);
                            if(val1.search(arr.valuetxt[1])>-1){
                                flag=true;
                                //     ss.getRange(1,1).setValue(x.toString()+y.toString());
                            }
                        }
                    }
                } 
                else {
                    if (arr.andor[0] == "Or") {
                                    ///////////////////////////////
                        var val1,ind1;
                        
                        if(arr.fieldoption[1]=="Due Date")
                            ind1=0;
                        if(arr.fieldoption[1]=="Label")
                            ind1=1;
                        if(arr.fieldoption[1]=="Last Person Connected")
                            ind1=3;
                        if(arr.fieldoption[1]=="Assignee")
                            ind1=5;
                        if(arr.fieldoption[1]=="Text Content")
                            ind1=4;
                        if(arr.fieldoption[1]=="Last Time")
                            ind1=2;
                        if(arr.fieldoption[1]=="List Name")
                            ind1=6;
                        val1=val[ind1];// 
                    
                        ////////////////////////
                    
                        if (arr.operatoroption[1] == "0") {
                            if (val1 == arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "1") {
                            if (val1 > arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "2") {
                            if (val1 < arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "3") {
                            if (val1 >= arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[1] == "4") {
                            if (val1 <= arr.valuetxt[1]) {
                                flag = true;
                            }
                        }
                        if(arr.fieldoption[1]!="Due Date"  && arr.operatoroption[1] !="0"){
                            if(val1.search(arr.valuetxt[1])>-1){
                            flag=true;
                            }
                        }
                    //ss.getRange(1,1).setValue(x.toString()+y.toString()+flag+arr.andor+arr.valuetxt[1]+val1);
                    }
                }


                if (flag) {
                    var val2,ind2;
                    
                    if(arr.fieldoption[2]=="Due Date")
                        ind2=0;
                    if(arr.fieldoption[2]=="Label")
                        ind2=1;
                    if(arr.fieldoption[2]=="Last Person Connected")
                        ind2=3;
                    if(arr.fieldoption[2]=="Assignee")
                        ind2=5;
                    if(arr.fieldoption[2]=="Text Content")
                        ind2=4;
                    if(arr.fieldoption[2]=="Last Time")
                        ind2=2;
                    if(arr.fieldoption[2]=="List Name")
                        ind2=6;
                    val2=val[ind2];
                    // ss.getRange(1,1).setValue(x.toString()+y.toString()+flag+arr.andor+val[2]);
                    ////////////////////////
                    // ss.getRange(1,1).setValue(arr.valuetxt[1]);
                    if (arr.andor[1] == "And") {
                        flag = false;
                    
                        if (arr.operatoroption[2] == "0") {
                            if (val2 == arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "1") {
                            if (val2 > con.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "2") {
                            if (val2 < arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "3") {
                            if (val2 >= con.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "4") {
                            if (val2 <= arr.valuetxt[2]) {
                                flag = true;
                            }
                        } 
                        if(arr.fieldoption[2]!="Due Date"){//  && arr.operatoroption[2] !="0"
                            // ss.getRange(1,1).setValue(val2);
                            if(val2.search(arr.valuetxt[2])>-1){
                                flag=true;
                                //     ss.getRange(1,1).setValue(x.toString()+y.toString());
                            }
                        }
                    }
                } 
                else {
                    if (arr.andor[1] == "Or") {
                                    ///////////////////////////////
                        var val2,ind2;
                        
                        if(arr.fieldoption[2]=="Due Date")
                            ind2=0;
                        if(arr.fieldoption[2]=="Label")
                            ind2=1;
                        if(arr.fieldoption[2]=="Last Person Connected")
                            ind2=3;
                        if(arr.fieldoption[2]=="Assignee")
                            ind2=5;
                        if(arr.fieldoption[2]=="Text Content")
                            ind2=4;
                        if(arr.fieldoption[2]=="Last Time")
                            ind2=2;
                        if(arr.fieldoption[2]=="List Name")
                            ind2=6;
                        val2=val[ind2];// 
                    
                        ////////////////////////
                    
                        if (arr.operatoroption[2] == "0") {
                            if (val2 == arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "1") {
                            if (val2 > arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "2") {
                            if (val2 < arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "3") {
                            if (val2 >= arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if (arr.operatoroption[2] == "4") {
                            if (val2 <= arr.valuetxt[2]) {
                                flag = true;
                            }
                        }
                        if(arr.fieldoption[2]!="Due Date"  && arr.operatoroption[2] !="0"){
                            if(val2.search(arr.valuetxt[2])>-1){
                            flag=true;
                            }
                        }
                    //ss.getRange(1,1).setValue(x.toString()+y.toString()+flag+arr.andor+arr.valuetxt[1]+val1);
                    }
                }




                if(flag){
                    var backcolor,bordercolor,textcolor;
                    backcolor=arr.backcolor;
                    bordercolor=arr.bordercolor;
                    textcolor=arr.textcolor;
                    if(arr.backcolor == "azure") backcolor="#007fff";if(arr.bordercolor=="azure") bordercolor="#007fff";if(arr.textcolor=="azure") textcolor="#007fff";
                    if(backcolor=="black") backcolor="#000000";if(bordercolor=="black") bordercolor="#000000";if(textcolor=="black") textcolor="#000000";
                    if(backcolor=="white") backcolor="#FFFFFF";if(bordercolor=="white") bordercolor="#FFFFFF";if(textcolor=="white") textcolor="#FFFFFF";
                    ss.getRange(x,y).setBackground(backcolor);
                    ss.getRange(x,y).setBorder(true, true, true, true, false, false, bordercolor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
                    ss.getRange(x,y).setFontColor(textcolor);
                    return;
                }
            }
        }
    }catch(e) {
        Logger.log(e);
    }
}

function deleteTrigger() {
    // Loop over all triggers.
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    // if (allTriggers[i].getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(allTriggers[i]);
    //     break;
    //  }
    }
}

var lists_all;
function getLists(key_and_token){
    var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/lists/all/?" + key_and_token);
    lists_all = JSON.parse((response.getContentText()));
}

var labels_all;
function getLabels(key_and_token){
    var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/labels/all/?" + key_and_token);
    labels_all = JSON.parse((response.getContentText()));
    return labels_all;
}

var members_all;
function getMembers(key_and_token){
    var response = UrlFetchApp.fetch(url + "boards/" + board_id + "/members/all/?" + key_and_token);
    members_all = JSON.parse((response.getContentText()));
    return members_all;
}

//called by google docs apps
var cardIDs=[];

function onOpen(e) {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .createMenu('Data Highlighter for Trello')
        .addItem('Start new session', 'Trello_Sheet_New')
        .addItem('Restore last session', 'Trello_Sheet_Restore')
    //     .addItem('4)Set Triggers',"createTimeDrivenTriggers")
        .addToUi();
//   deleteTrigger();
}

function Trello_Sheet_New() {
    var html = HtmlService.createHtmlOutputFromFile('new')
      .setTitle('Data Highlighter for Trello')
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function Trello_Sheet_Restore() {
    var html = HtmlService.createHtmlOutputFromFile('restore')
        .setTitle('Data Highlighter for Trello')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

var today;
var todayDay;
var todayDate;

//var ui = SpreadsheetApp.getUi(); 
//var sh = SpreadsheetApp.getActiveSheet();
//var rg = sh.getDataRange();
//var vA = rg.getValues();
//var rowsFound = 0;
//var rowsDeleted = 0;

//function ClearOld(value) {
//  var clearDate=value;
//  var testThis = ui.alert(
// 'Clear entries prior to',
// clearDate,
//  ui.ButtonSet.YES_NO);
//}
//function onOpen() {
//    var ui = SpreadsheetApp.getUi();
////    var mainMenu = ui.createMenu("Orientation");
////    mainMenu.addItem("Flip Selection Vertically", "flipItMenuVertical");
////    mainMenu.addSeparator();
////    mainMenu.addItem("Flip Selection Horizontally", "flipItMenuHorizontal");
////    mainMenu.addToUi();
//    ui.showSidebar(createUI());
//};
// 
//function createUI() {
//    var app = UiApp.createApplication();
//    app.setTitle('Flip Selection');
// 
//    var radioFlipVertical = app.createRadioButton('radioFlip').setId('radioFlipVertical').setText('Flip Selection Vertically.');
//    var radioFlipHorizontal = app.createRadioButton('radioFlip').setId('radioFlipHorizontal').setText('Flip Selection Horizontally.');
//    var buttonFlip = app.createButton("Flip Selection");
//    var panel = app.createVerticalPanel();
// 
//    panel.add(radioFlipVertical);
//    panel.add(radioFlipHorizontal);
//    panel.add(buttonFlip);
//    app.add(panel);
// 
//    var handlerRadioButtons = app.createServerHandler('radioButtonsChange');
//    radioFlipVertical.addValueChangeHandler(handlerRadioButtons);
//    radioFlipHorizontal.addValueChangeHandler(handlerRadioButtons);
// 
//    var handlerButton = app.createServerHandler('flipIt');
//    buttonFlip.addClickHandler(handlerButton);
// 
//    return app;
//}
// 
//function flipItMenuVertical() {
//    var ui = SpreadsheetApp.getUi();
//    var result = ui.prompt('Destination', 'Please enter the column name to start in e.g. C :', ui.ButtonSet.OK_CANCEL);
//    var button = result.getSelectedButton();
//    var targetCell = result.getResponseText();
//    if (button == ui.Button.OK) {
//        flipSelection('vertical', targetCell);
//    }
//}
// 
//function flipItMenuHorizontal() {
//    var ui = SpreadsheetApp.getUi();
//    var result = ui.prompt('Destination', 'Please enter the column name to start in e.g. C :', ui.ButtonSet.OK_CANCEL);
//    var button = result.getSelectedButton();
//    var targetCell = result.getResponseText();
//    if (button == ui.Button.OK) {
//        flipSelection('horizontal', targetCell);
//    }
//}
// 
//function radioButtonsChange(e) {
//    ScriptProperties.setProperty('selectedRadio', e.parameter.source);
//}
// 
//function flipIt(e) {
//    var ui = SpreadsheetApp.getUi();
// 
//    var result = ui.prompt('Destination', 'Please enter the column name to start in e.g. C :', ui.ButtonSet.OK_CANCEL);
//    var button = result.getSelectedButton();
//    var targetCell = result.getResponseText();
//    if (button == ui.Button.OK) {
//        var selectedOption = ScriptProperties.getProperty('selectedRadio');
//        if (selectedOption == 'radioFlipVertical') {
//            flipSelection('vertical', targetCell);
//        } else {
//            flipSelection('horizontal', targetCell);
//        }
//    }
//}
// 
//function flipSelection(orientation, target) {
//    var sheet = SpreadsheetApp.getActiveSheet();
//    var selectedValues = sheet.getActiveRange().getValues();
//    var range;
//    var startColIndex = sheet.getRange(target + '1').getColumn();
// 
//    sheet.getActiveRange().clear();
//    if (orientation == "horizontal") {
//        for (var i = 0; i < selectedValues.length; i++) {
//            range = sheet.getRange(1, startColIndex + i);
//            range.setValue(String(selectedValues[i]));
//        }
//    } else if (orientation == "vertical") {
//        var vals = String(selectedValues[0]).split(",");
//        var rowCount = 1;
//        vals.forEach(function (value) {
//            range = sheet.getRange(rowCount, startColIndex);
//            range.setValue(value);
//            rowCount++;
//        });
//    }
//}