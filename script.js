var siteURL = "https://milliman.sharepoint.com/sites/LTS/mgalfa/dev/";
var imgLibraryExtension = "Documents/Sharepoint%20Banner/images/";
//https://milliman.sharepoint.com/sites/LTS/mgalfa/dev/Documents/Sharepoint%20Banner/banner.html


//
ExecuteOrDelayUntilScriptLoaded(retrieveRecentListItems, "sp.js");

function queryforItems() {
    for (i = 0; i < listsToShow.length; i++) {
        var currentList = listsToShow[i];
        lists[i] = ctx.get_web().get_lists().getByTitle(currentList);
        if (currentList == 'Calendar') {
            items[i] = lists[i].getItems(camlQueryforCalendar);
        } else if (currentList == 'Documents') {
            items[i] = lists[i].getItems(camlQueryforDocuments);
        } else {
            items[i] = lists[i].getItems(camlQueryforTasks);
        }
        ctx.load(lists[i]);
        ctx.load(items[i]);
    }
}
function getDocumentTitle(currentItem) {
    var itemTitle = currentItem.get_item('FileLeafRef');
    var URL = currentItem.get_item('FileRef');
    var title = "<a href='" + URL + "'>" + itemTitle + "</a>";

    return title;
}
function getTaskTitle(currentItem) {
    var URL = siteURL + "Lists/Product%20Development%20Team%20Tasks/DispForm.aspx?ID=" + currentItem.get_item('ID');
    var itemTitle = currentItem.get_item('Title');
    var title = "<a href='" + URL + "'>" + itemTitle + "</a>";

    return title;
}
function adjustForTimeZone(date) {
    date.setHours(date.getHours() + 8);
    return date;
}
function getCalendarData(currentItem) {
    var data = {
        "startDate": adjustForTimeZone(currentItem.get_item('EventDate')).toDateString(),
        "endDate": adjustForTimeZone(currentItem.get_item('EndDate')).toDateString(),
        "URL": siteURL + "Lists/Calendar/DispForm.aspx?ID=" + currentItem.get_item('ID')
    };

    return data;
}
var outputOtherData = function (title, modified, author, itemNumber) {
    console.log("title: " + title + ", modified on " + modified.toString() + " by " + author);
    var newCap = "cap" + itemNumber;
    var divForCap = "<div id='" + newCap + "'></div>";
    var list = listsToShow[i].replace(/\s/g, '');

    $(divForCap).appendTo("#" + list + ' > div');
    //if (listsToShow[i] == 'Documents') {
    //    var docIcon = "<span class='image'><img src='https://milliman.sharepoint.com/sites/LTS/mgalfa/dev/Documents/Sharepoint%20Banner/images/documenticon.png' /><br></span>";
    //    $(docIcon).appendTo("#" + newCap);
    //}
    $("<span class='title'>" + title + "</span><br>").appendTo("#" + newCap);
    $("<span class='author'>" + author + "<span>").appendTo("#" + newCap);
    $("<span class='modified'>" + modified.toString() + "</span><br>").appendTo("#" + newCap);

    //document.getElementById('cap' + (itemNumber).toString()).innerHTML =
    //    docIcon   
    //    + "<span class='title'>" + title + "</span>"
    //    + "<span class='author'>" + author + "</span>"
    //    + "<span class='modified'>" + modified.toString() + "</span>";
}
var outputCalData = function (author, calData, j) {
    if (new Date <= endDate && new Date >= startDate) {
        var linkedDate = "<a href='" + calData.URL + "'>" + calData.endDate + "</a>";
        $("<div>" + author + " until " + linkedDate + "</div>").appendTo("#Calendar #gone div");
    } else if (new Date < startDate) {
        var linkedDate = "<a href='" + calData.URL + "'>" + calData.startDate + " -> " + calData.endDate + "</a>";
        $("<div>" + author + " from " + linkedDate + "</div>").appendTo("#Calendar #willBeGone div");
    }
}
function outputData() {
    var j = 0;
    for (i = 0; i < items.length; i++) {
        console.log("'" + listsToShow[i] + " List/Library' is about to enumerate");
        var listItemEnumerator = items[i].getEnumerator();

        while (listItemEnumerator.moveNext()) {
            j++;
            var currentItem = listItemEnumerator.get_current();
            var modified = currentItem.get_item('Modified').toDateString();
            var author = currentItem.get_item('Author').get_lookupValue();

            if (listsToShow[i] == "Documents") {
                var title = getDocumentTitle(currentItem);
                outputOtherData(title, modified, author, j);
            }
            else if (listsToShow[i] == "Calendar") {
                var calData = getCalendarData(currentItem);
                outputCalData(author, calData, j);
            }
            else {
                var title = getTaskTitle(currentItem);
                outputOtherData(title, modified, author, j);
            }
        }
    }
}
function interactiveSlideshow(){
    var contentElements = $("#ProductDevelopmentTeamTasks div > div");
    var k = contentElements.length - 1;
    var terminate = false;
    var delay = 7000;
    $(contentElements).hide();

    var loop = function (){
        $(contentElements[k]).fadeOut([100]).queue(
          function (){
              k++;
              k = k % contentElements.length;
              $(contentElements[k]).fadeIn([100]);
              $(this).dequeue();
          })

        if (terminate){
            return;
        }
        timeout = setTimeout(loop, delay);
    };

    loop();

    $("#next").click(function () {
        clearTimeout(timeout);
        terminate = true;
        contentElements.fadeOut(100);
        terminate = false;
        loop();
    });

    $("#previous").click(function () {
        clearTimeout(timeout);
        terminate = true;
        contentElements.fadeOut(100);
        if (k == 0) {
            k = contentElements.length - 2;
        }
        else {
            k--;
            k--;
        }
        terminate = false;
        loop();
    });
}
function onQuerySucceeded(sender, args) {
    outputData();
    interactiveSlideshow();
}
function onQueryFailed(sender, args) {
    console.log("Query Failed" + args.get_message() + '\n' + args.get_stackTrace());
}

function retrieveRecentListItems() {
    var camlQueryforDocuments = new SP.CamlQuery();
    camlQueryforDocuments.set_viewXml("<View Scope='Recursive'><Query><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy></Query><RowLimit>10</RowLimit></View>");

    var camlQueryforCalendar = new SP.CamlQuery();
    camlQueryforCalendar.set_viewXml("<View Scope='Recursive'><Query><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy></Query><RowLimit>2</RowLimit></View>");

    var camlQueryforTasks= new SP.CamlQuery;
    camlQueryforTasks.set_viewXml("<View Scope='Recursive'><Query><OrderBy><FieldRef Name='Modified' Ascending='False' /></OrderBy></Query><RowLimit>5</RowLimit></View>");

    var ctx = new SP.ClientContext.get_current();
    var listsToShow = ['Documents', 'Product Development Team Tasks', 'Calendar'];
    var lists = [];
    var items = [];
    queryforItems;
    ctx.executeQueryAsync(onQuerySucceeded, onQueryFailed);
}

