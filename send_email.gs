function sendEmail() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var last_row = sheet.getLastRow()
  var po_num = sheet.getName()
  
  var item_list = []
  for(var i = 2; i <= last_row; i++) {
    var qty_recd = sheet.getRange(i, 7).getValue()
    var qty_notd = sheet.getRange(i, 8).getValue()
    var qty_expd = sheet.getRange(i, 5).getValue()
    var completed_box = sheet.getRange(i, 9).getValue()
    if(completed_box == false) {
      if(qty_recd > qty_notd) {
        var line_item ="<br>" + sheet.getRange(i, 1).getValue() + " x <b>" + qty_recd +" out of " + qty_expd + " " + sheet.getRange(i, 3).getValue() + "</b>";
        item_list.push(line_item)
        sheet.getRange(i, 8).setValue(qty_recd)
        var updated_qty_notd = sheet.getRange(i, 8).getValue()
        if(updated_qty_notd == qty_expd) {
          sheet.getRange(i, 9).setValue("TRUE")
        }
      }
    }
  }
  
  if (item_list.length == 0) {
    SpreadsheetApp.getUi().alert("There are no items to send.")
    return
  };
  
  order_completion = ""
  for(var i = 2; i <= last_row; i++) {
    var completed_box = sheet.getRange(i, 9).getValue()
    if(completed_box == false) {
     order_completion = ""
     break
    } else if(completed_box == true) {
       order_completion = "\n\nThere are no further items expected from this order"
    }
  }
    
  var bodyHTML = "<body><p>Hello,</p><p>I have received the following items from the PO# listed in the subject line:</p><p>" + item_list + "</p><br><p>" + order_completion + "</p><p>Thank you,<br>Logan Severin</p>"
  for (var i=0; i<item_list.length; i++) {
    item_list[i] = item_list[i].replace(/<[^>]+>/ig, "")
    item_list[i] = "\n\n" + item_list[i]
  }
  var body = "Hello, \n\nI have received the following items from the PO# listed in the subject line:" + item_list + order_completion + "\n\nThank you,\nLogan"
  
  
  MailApp.sendEmail("interested_party@somedomain.com", po_num, body, { htmlBody: bodyHTML});
  
}