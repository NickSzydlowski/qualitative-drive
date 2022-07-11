//This function extracts the comments from a set of Google Documents and write them to a Google Spreadsheet. It is intended to facilitate qualitative coding in Google Drive.

function listComments() {
  var fileList=[];
  //Get list of files in the folder
  var files = Drive.Files.list({ 
      "corpora": "drive",
      "driveId": "",//Shared Drive ID
      "includeTeamDriveItems": true,
      "q": "" in parents", //Folder ID must appear as a parent
      "supportsTeamDrives": true
    })
   if (files.items && files.items.length > 0) {
     for (var b = 0; b < files.items.length; b++) {
       var file=files.items[b]
       if (file.explicitlyTrashed) {//Ignore deleted files
       }
       else {
          var mType = file.mimeType;
          if ( mType === 'application/vnd.google-apps.document') { //Only look at Google Docs
            fileList.unshift(file.id) //Add file ID to list
          }
       }
   }
}
//Create empty arrays for values
var hList = [], cList = [], aList = [], bList = [], dList = [], eList=[], fList=[], gList=[]; jList=[];
//Count is an overall count of rows
var count=1;

//Loop through the files on the list
for (var z = 0; z < fileList.length; z++) {  
  var token=true;//We're using token to detect when there are no more comments on a given file
  var arg = {
    maxResults: 100
  }; 

while (token) {//Loops through the different pages of comments from a single document - if there are more than 100 comments, there will be multiple pages
console.log(fileList[z]);//Helpful to watch the progress - you will see file IDs in the console
  var comments = Drive.Comments.list(fileList[z], arg); //Get the comments
   if (comments.items && comments.items.length > 0) {
    for (var i = 0; i < comments.items.length; i++) { //Loop through the comments
      var comment = comments.items[i]; 
      //If the comment has line breaks, it represents more than one code. The script writes each code as a separate row.
       var splitComments = comment.content.split(/\r\n|\r|\n/g); 
       
      for (var n=0; n<splitComments.length; n++) { //Loop through the separate codes. Usually there is only one, but if there are more, we want them on separate rows

      if (comment.context && comment.context.value !== 'undefined') { //The context was throwing errors - I did not figure out why
        hList.unshift([comment.context.value]); //Write the context, which is the highlighted document text for the comment
      }
      else {
        hList.unshift([' '])
      };
      cList.unshift([splitComments[n]]);//Write the comment text, split at the line break
      aList.unshift([comment.fileId]);//Write the file ID
      bList.unshift([comment.fileTitle]);//Write the file name
      dList.unshift([comment.author.displayName]);//Write the commenter's name
      eList.unshift([count]);//Write the record number
      count=count+1;//Advance the record number
      fList.unshift(["comment"]);//This item is an original comment
      gList.unshift([comment.commentId])//The comment ID - for bringing together comments and replies if needed
      jList.unshift([comment.deleted]);//Was the comment deleted?
      
      if (comment.replies.length>0) {//if the comment has replies write each to a new row

       for (var k = 0; k < comment.replies.length; k++) {//Loop through the comment's replies - the fields are the same as above
      var splitReplies = comment.replies[k].content.split(/\r\n|\r|\n/g); 
       
      for (var p=0; p<splitReplies.length; p++) { //Loop through the separate codes. Usually there is only one, but if there are more, we want them on separate rows
        hList.unshift([comment.context.value]);
        cList.unshift([splitReplies[p]]);
        aList.unshift([comment.fileId]);
        bList.unshift([comment.fileTitle]);
        dList.unshift([comment.replies[k].author.displayName]);
        eList.unshift([count]);
        count=count+1;
        fList.unshift(["reply"]);//This item is a reply
        gList.unshift([comment.commentId]);
        jList.unshift([comment.replies[k].deleted]);
      }//end of loop for reply text split on the line break 
      }//end of replies loop
      }//end of replies if statement
    }//end of loop for comment text split on the line break
    }//end of loop for comments
    }//end of comments if statement
   
  if (comments.nextPageToken) {//if there is a next page of comments - the maximum number per request is 100
    token=true;
    var arg = {//These arguments are for a new request using the token from the previous request
      maxResults: 100,
      pageToken: comments.nextPageToken
    };
  }//end of token if statement
  else {//If there is no token, end the while loop and advance to the next file
    token=false;
  }
}//end of token while loop
}//end of files loop

// Write the values from our lists to the spreadsheet
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange("A1:A" + hList.length).setValues(hList);
    sheet.getRange("B1:B" + cList.length).setValues(cList);
    sheet.getRange("C1:C" + aList.length).setValues(aList);
    sheet.getRange("D1:D" + bList.length).setValues(bList);
    sheet.getRange("E1:E" + dList.length).setValues(dList);
    sheet.getRange("F1:F" + eList.length).setValues(eList);
    sheet.getRange("G1:G" + fList.length).setValues(fList);
    sheet.getRange("H1:H" + gList.length).setValues(gList);
    sheet.getRange("I1:I" + jList.length).setValues(jList);
}

