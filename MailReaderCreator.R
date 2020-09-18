#MailReader.R
# v 202008251022
# Script that pre-elaborates mails and
#  uniteresting ones: set to "read"
#  transactions
#   download mail and attachments and upload to SP
#   flag planner task


######################## script: please do not change ########################
##packages
library(AzureGraph)
library(AzureAuth)
library(jsonlite)
library(tidyverse)
library(bizdays)
library(httr)
library(functional)
library(gtools)
library (stringi)
library(httpuv)
library(collections)


#scopes
# needed for the login; indicates the permissions
scopes <- c("https://graph.microsoft.com/User.Read",
            "https://graph.microsoft.com/Mail.ReadWrite",
            "offline_access")

#TX_MARKER
# regex pattern that recognizes TX ids
# the pattern is first TX, then another uppercase letter, then 6 digits
TX_MARKER = "TX[A-Z]\\d{6}"

user_assignments_names = list(c("gazzeal", "AG"),
                           c("nikolma", "MN"),
                           c("papadde", "DP"),
                           c("gazzealefsa", "AG"),
                           c("garettba", "BG"),
                           c("salf", "AG"),
                           c("vargazo", "ZV"))

user_list = list(user_names = c('AG', 'MN', 'DP', 'ZV', 'BG', 'unassigned'),
                 user_ids = c('3963a97a-43e7-43e7-b71a-a710b7bbc4fa','7900481f-40d8-4e73-868b-6db453b53aa4','0db6b33a-4615-43af-85de-cab76eb2628c','37a32389-1aac-481c-a1fc-e80ab2e5aa34', 'ab4bbfc2-6d90-44f6-b832-86ae7c6464a2', 'NA'))
user_tib = as_tibble(user_list)

##global objects
#execute_login
# performs the login and gets permissions; this cannot be done in a function
#get login token using AzureAuth for reading emails
# exact permissions are defined above in scopes
token=AzureAuth::get_azure_token(scopes,tenant = "406a174b-e315-48bd-aa0a-cdaddc44250b","970559dd-4ac3-488f-a7fe-9cdf47a66233",version = 2, auth_type="device_code", password="jbNLB_br7iBFoNB-_~o35Gd_35fvbxP72S", use_cache = TRUE) 
print("going to refresh azureauth token")
token$refresh()
print("refreshed azureauth token; going to request azuregraph token")
#get login token using AzureGraph for planner and sharepoint
if (!exists("me")) {
  gr=create_graph_login(auth_type="device_code")
  me <- gr$get_user()
}

my_teams = call_graph_url(me$token,
                          "https://graph.microsoft.com/beta/me/joinedTeams")

##functions
#flags for counting how many mails are flagged read
#read comment under main function for an explanationo of the need
amount_set_read = 0
read_flag_was_set = FALSE
PACKAGE_SIZE = 10

#main
# reads using the graph api one packet of unread mails at a time and calls the elaboration function
# for paging, instead of using the link provided by graph in odata.nextLink we are obliged to 
# compute the number of skipped mails manually, because the nextLink does not take into account
# that we have set the read flag for some mail. So some mails would be ignored using nextLink
# this is a typical example of iteration over a set that is being changed while you iterate.
# example: if the first package is loaded, mails number 11, 12 and 13 are not in the first package;
#          assuming e.g. that 3 of the first 10 mails are set read, when we call nextLink to load the 
#          second package (and it will skip the first 10 unread mails to avoid the first package)
#          we will skip also number 11, 12 and 13 because the set of unread emails has changed
main = function() {
  amount_elaborated = 0 #counts how many mails have undergone elaboration, independently of read flags set
  #read my messages (only the first 10) 
  my_messages = AzureGraph::call_graph_endpoint(token,paste("me/mailFolders/inbox/messages?$filter=isRead+eq+false&$orderby=sentDateTime",sep=""))
  my_messages_value = my_messages$value
  elaborate_mail_package(my_messages_value)
  amount_elaborated = amount_elaborated + PACKAGE_SIZE
  while(!is.null(my_messages$'@odata.nextLink')) {
    print(paste("In main; amount elaborated: ", amount_elaborated, "; amount set read: ", amount_set_read))
    my_messages = AzureGraph::call_graph_endpoint(token,paste("me/mailFolders/inbox/messages?$filter=isRead+eq+false&$orderby=sentDateTime&$skip=",amount_elaborated - amount_set_read,sep=""))
    my_messages_value = my_messages$value
    elaborate_mail_package(my_messages_value)
    amount_elaborated = amount_elaborated + PACKAGE_SIZE
  }
}

#main_step_by_step
# performs the elaboration only of the first mail
main_step_by_step = function() {
  #read my messages (only the first 10) 
  my_messages = AzureGraph::call_graph_endpoint(token,paste("me/mailFolders/inbox/messages?$filter=isRead+eq+false&$orderby=sentDateTime",sep=""))
  my_messages_value = my_messages$value
  elaborate_mail_package(list(my_messages_value[[1]]))
}

#main iteration over the messages
elaborate_mail_package = function (my_messages_value) {
  sapply(my_messages_value, elaborate_mail)
}
   
#read csv file with parameter information
read_parameter_table = function(fileName, teamName) {
  #search for my teams
  TSTA_team = my_teams$value[lapply(my_teams$value, function(x) x$displayName) %in% teamName == TRUE]
  TSTA_team_id = TSTA_team[[1]]$id
  
  # id of the team transformation assurance 58398e3c-899e-431f-99d4-207f15012489
  #call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/58398e3c-899e-431f-99d4-207f15012489/drive"))
  group_drive = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/",TSTA_team_id,"/drive", sep=""))
  group_drive_id = group_drive$id
  
  search_csv = call_graph_url(me$token,
                              paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/root/search(q='{",fileName,"}')", sep="")) 
  csv_id = search_csv$value[[1]]$id
  
  #read the file
  csv_param = call_graph_url(me$token,paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/items/",csv_id,"/content", sep=""))
  
  # convert a sequence of unicode hex to a vector or charachters 
  csvchars = chr(csv_param)
  
  # convert a vector of charachters to a string
  csvtable = paste(csvchars, collapse="")
  csvLongStr = toString (csvtable)
  csv_fine_str = stri_enc_toutf8(stri_trans_nfd(csvLongStr)) #this final command eliminates the special charachters
  
  # create a tibble from the csv data read from the file and converted using the previous snippets
  read_csv(csv_fine_str)
}

plans_and_groups = read_parameter_table("PlanGroupData.csv", "Transformation Assurance")

#detect if a mail is a taskify request
detect_taskify = function(curr_msg){
  grepl(paste("^", "Taskify", sep=""), curr_msg$subject, ignore.case = TRUE)
}

#detect if a mail is restricted and thus shall not be published
detect_restricted = function(curr_msg) {
  grepl("Restricted", curr_msg$subject, ignore.case = TRUE) ||
    grepl("confidential", curr_msg$subject, ignore.case = TRUE)
}

#elaborate_mail
# elaboration of one single mail
# a series of conditions is used to decide if
# mail shall be linked to a transaction (see elaborate_tx)
# or simpli flaged read
elaborate_mail = function(curr_msg) {
  read_flag_was_set <<- FALSE
  message_id = curr_msg$id  
  print(paste("Starting elaboration of mail with subject: ", curr_msg$subject,sep=""))

  must_read = FALSE
  shall_be_read = TRUE
  
  #taskify is detected always, independently of all other criteria
  if (detect_taskify(curr_msg)) {
    taskify(curr_msg)
    shall_be_read = FALSE
  }    
  
  #from VIP has to be read
  if (must_read(curr_msg)) {
    must_read = TRUE
    shall_be_read = TRUE
  }

  #identify automated mails 
  if ( SET_READ_AUTOMATED_MESSAGES && detect_automated_message (curr_msg)) {
    print(paste("it is an automated message with sender ", curr_msg$from$emailAddress$address))
    shall_be_read = FALSE
  }
  
  #exclude mails by subject
  if (SET_READ_BY_SUBJECT && detect_optional_by_subject (curr_msg)) {
    print(paste("it is a message with subject ", curr_msg$subject))
    shall_be_read = FALSE
  }
  
  #invitations
  #location exists 
  if (SET_READ_INVITATIONS && detect_invitation_reply (curr_msg)) {
    print(paste("it is an invitation with subject ", curr_msg$location, curr_msg$startDateTime))
    shall_be_read = FALSE
  }
  
  #planner
  if (SET_READ_PLANNER && detect_planner (curr_msg)) {
    print(paste("it is a planner message with subject ", curr_msg$subject))
    shall_be_read = FALSE
  }
      
  #servicenow except hold
  if (SET_READ_SERVICE_NOW_EXCEPT_HOLD && detect_service_now_not_hold (curr_msg)) {
    print(paste("it is a service now message not hold with subject ", curr_msg$subject))
    shall_be_read = FALSE
  }

  #servicenow all
  if (SET_READ_SERVICE_NOW_ALL && 
      (detect_service_now_hold (curr_msg) ||
       detect_service_now_not_hold(curr_msg))) {
    print(paste("it is a service now message with subject ", curr_msg$subject))
    shall_be_read = FALSE
  }

  #elaborate tx shall only start if the previous criteria do not exclude the mail
  if (shall_be_read) {
    #identify tx messages and elaborate them 
    if (ELABORATE_TX_MAILS && detect_tx_mail(curr_msg) && !detect_restricted(curr_msg)) {
      elaborate_tx(curr_msg)
    }
  }

  #the following criteria do not block elaborate_tx
  
  #pure cc
  #to does not contain gazzetta nor team ts ta services
  if (SET_READ_ALL_CC && detect_pure_cc(curr_msg)) {
    print(paste("it is a pure cc message with recipient ", curr_msg$toRecipients))
    shall_be_read = FALSE
  }  

  #if it is not a must read and a rule fired to ignore it, mark the mail as already read and
  #increase the number of mails flagged read; see the main function for the explanation of the need of this
  print(paste("In elaborate mail; must_read is ",must_read," and shall_be_read is: ", shall_be_read, "; and amount_set_read is: ", amount_set_read))
  if (!must_read && !shall_be_read) {
    set_read_flag(curr_msg)
    amount_set_read <<- amount_set_read + 1
    print(paste("In elaborate mail; increased amount_set_read to ", amount_set_read))
  }
}

#must_read
must_read = function(curr_msg) {
  detect_VIP_message(curr_msg) ||
    curr_msg$importance == "high"
}

#detect_pure_cc
detect_pure_cc = function(curr_msg) {
  current_user_mail = me$properties[["mail"]]
  to_groups = append(to_groups, current_user_mail)
  match_list = sapply(to_groups, function(x) find_recipient(curr_msg , x))
  pure_to = TRUE %in% match_list
  !pure_to
}

#find_recipient
find_recipient = function (curr_msg, recipient) {
  match_list = sapply(curr_msg$toRecipients, function (x) grepl(recipient, x, ignore.case = TRUE))
  TRUE %in% match_list
}

#detect_optional_by_subject
detect_optional_by_subject = function(curr_msg) {
  match_list = sapply(exclude_subjects_beginning_with, function (x) grepl(paste("^", x, sep=""), curr_msg$subject, ignore.case = TRUE))
  match_list = append(match_list, 
                      sapply(exclude_subjects_containing, function (x) grepl(x, curr_msg$subject, ignore.case = TRUE)))
  match_list = append(match_list, 
                      sapply(exclude_subjects_ending_with, function (x) grepl(paste(x, "$", sep=""), curr_msg$subject, ignore.case = TRUE)))
  
  TRUE %in% match_list
}

#detect_service_now_hold
detect_service_now_hold = function(curr_msg) {
  !grepl("^Your case is on hold waiting for your input", curr_msg$subject) &&
    curr_msg$from$emailAddress$address == "EFSAProcurement@efsa.europa.eu"
}

#detect_service_now_not_hold
detect_service_now_not_hold = function(curr_msg) {
  (grepl("^Your case is on hold waiting for your input", curr_msg$subject) ||
    grepl("^A new case has been registered", curr_msg$subject) ||
     grepl("^Your request has received an update", curr_msg$subject) ||
     grepl("^A new Commitment approval has been submitted to the AO", curr_msg$subject) ||
     grepl("^Your Commitment has been approved by the AO", curr_msg$subject) ||
     grepl("^The contract has been signed", curr_msg$subject) ||
     grepl("^A new contract has been sent for signature to the supplier", curr_msg$subject)) &&
    curr_msg$from$emailAddress$address == "EFSAProcurement@efsa.europa.eu"
}

#detect_automated_message
# from efsa.jiveon.com; noreply@Planner.Office365.com; ec-fp-internet-services-do-not-reply@ec.europa.eu; noreply@efsa.europa.eu; noreply@yammer.com; noreply@sciforma.net; u4u-news-com-request@lists.u4unity.eu; REP-PERS-OSP-U4U@ec.europa.eu; noreply@email.teams.microsoft.com
detect_automated_message = function(curr_msg) {
  grepl(curr_msg$from$emailAddress$address, automated_senders, ignore.case = TRUE)
}

#detect_VIP_message
detect_VIP_message = function(curr_msg) {
  grepl(curr_msg$from$emailAddress$address, VIP_senders, ignore.case = TRUE)
}

#detect_own_message
#own messages do not flag the task as having a new message
detect_own_message = function(curr_msg) {
  TRUE %in% grepl(tolower(curr_msg$from$emailAddress$address), tsta_member_mails, ignore.case = TRUE)
}

#detect_invitation_reply
detect_invitation_reply = function(curr_msg) {
  !is.null(curr_msg$location) ||
    !is.null(curr_msg$startDateTime)
}

#detect_planner
detect_planner = function(curr_msg) {
  grepl("^Your request is on hold", curr_msg$subject) ||
    grepl("^You have late tasks", curr_msg$subject) ||
    grepl("^Comments on task", curr_msg$subject)
}

#move_messsage
# function that moves a message identified by the id
# to a folder identified by its id
move_message = function (msg_id, dest_folder_id) {
  body_obj = list (destinationId = dest_folder_id)
  body_json = toJSON (body_obj, auto_unbox = TRUE)
  response = call_graph_url(token,paste("https://graph.microsoft.com/beta/me/messages/",curr_msg$id,"/move" ,sep=""),  #we use beta because it can handle the priority field
                            body=body_json, http_verb=c ("POST"),encode = "raw")  
  response  
}

#search_tx_drive
# searches the appropriate tx drive
search_tx_drive = function(name_component, team_name) {
  #search for my teams
  myTeams = call_graph_url(me$token,"https://graph.microsoft.com/v1.0/me/joinedTeams")
  TSTA_team = myTeams$value[lapply(myTeams$value, function(x) x$displayName) %in% team_name == TRUE]
  TSTA_team_id = TSTA_team[[1]]$id
  
  # id of the team transformation assurance 58398e3c-899e-431f-99d4-207f15012489
  #call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/58398e3c-899e-431f-99d4-207f15012489/drive"))
  group_drive = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/",TSTA_team_id,"/drive", sep=""))
  group_drive_id = group_drive$id
}

#search_tx_folder
# searches the appropriate tx folder
# needs a team name to focus the search
search_tx_folder = function(name_component, group_drive_id) {
  search_result = call_graph_url(me$token,
                                 paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/root/search(q='{",name_component,"}')", sep="")) 
  result_id = search_result$value[[which(sapply(search_result$value, function(x) !is.null(x$folder)))]]$id
  result_id
}

#search_mail_in_folder
# searches if the mail message has already been
# saved to the indicated folder
search_mail_in_folder = function(curr_msg, folder_id, group_drive_id) {
  #get list of files under the folder
  folder_files = call_graph_url(me$token, paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/items/",folder_id,"/children", sep=""))
  
  if (length(folder_files$value) > 0) {
    #we compare only the first 13 chars that contain the date and exact time; so the file can be renamed and found anyway
    saved_file = which(sapply(folder_files$value, function(x) grepl(substr(generate_file_name(curr_msg), 1, 16), substr(x$name, 1, 16),fixed=TRUE)))
    num_found = length(saved_file)
    
    if (num_found != 0) {
      folder_files$value[[saved_file]]
    } else {
      NULL
    }
  } else {
    NULL
  }
}

#extract name of sender
extract_name_of_sender = function (curr_msg) {
  sender = curr_msg$from$emailAddress
  if (sender$name == "") {
    name_with_at = regmatches(sender$address, regexpr("[a-z|A-Z|1-9|\\.]*\\@", sender$address))
    substring(name_with_at, 1, nchar(name_with_at) - 1)
  } else {
    str_replace_all(sender$name, "[^[A-Z|a-z] ]", "") 
  }
}

#generate_file_name
#todo: use the sent instead of the systime; who cares of the time when you run the script; sent date time is relevant
generate_file_name = function (curr_msg, tx_id = "") {
  if (tx_id != "") {
    paste(format(strptime(curr_msg$sentDateTime, format="%Y-%m-%dT%H:%M:%SZ") + 7200, "%Y%m%d %H%M%S"), " ", extract_name_of_sender(curr_msg), " ", substring(str_replace_all(curr_msg$subject, "[^[:alnum:]]", " "), 1, 30), " (", tx_id, ")", ".eml", sep="")
  } else {
    paste(format(strptime(curr_msg$sentDateTime, format="%Y-%m-%dT%H:%M:%SZ") + 7200, "%Y%m%d %H%M%S"), " ", extract_name_of_sender(curr_msg), " ", substring(str_replace_all(curr_msg$subject, "[^[:alnum:]]", " "), 1, 30), " (", detect_tx(curr_msg$subject), ")", ".eml", sep="")
  }
}

#generate_folder_name
generate_folder_name = function (curr_msg) {
  paste(format(strptime(curr_msg$sentDateTime, format="%Y-%m-%dT%H:%M:%SZ") + 7200, "%Y%m%d %H%M%S"), "-Attachments", sep="")
}

#write_mail
# function that writes a mail read using the graph api to 
#  a file on the hd ready to be uploaded
#  does not save attachments
#  .eml is produced "manually"
#  works only for html mails for now
# sources
#  https://mailchannels.zendesk.com/hc/en-us/articles/360005491712-How-to-Create-an-EML-File
#  https://docs.microsoft.com/en-us/graph/outlook-get-mime-message
#  https://docs.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0
write_mail = function (curr_msg, tx_id = "") {
  #get mime version of the message
  mime_message = AzureGraph::call_graph_endpoint(token,
                                                 paste("me/messages/", 
                                                       curr_msg$id, 
                                                       "/$value",
                                                       sep=""))
  
  file_name = generate_file_name(curr_msg, tx_id = tx_id)

  #write the MIME message manually
  write(gsub("\r\n", "\n", mime_message), file_name)
}

#delete_temp_mail_file
# deletes the temporary copy of the mail file
delete_temp_mail_file = function(curr_msg, tx_id = "") {
  file_name = generate_file_name(curr_msg, tx_id = tx_id)
  if (file.exists(file_name)) file.remove(file_name)  
}

#delete_temp_attachment_files
# deletes the temporary copies of the attachment files
delete_temp_attachment_files = function(attachment_names_and_types) {
  sapply(attachment_names_and_types, function(x) {if(file.exists(x[[1]])) file.remove(x[[1]])})
  
}

#save attachments
# use the id to retrieve the list of attachments
# the content of the attachment is already in the list
# no need to call separately get for each attachment
# by id as suggested on the web
save_attachments = function(curr_msg) {
  attachment_names_and_types = list()
    message_attachments = AzureGraph::call_graph_endpoint(token,
                                                        paste("me/messages/", 
                                                              curr_msg$id, 
                                                              "/attachments",
                                                              sep=""))
  
  count = 1
  for (att in message_attachments$value) {
    #write the attachment to the hard disc
    #goes to Documents of the current user
    
    if(!is.null(att$contentBytes)) { #some attachments arrive empty, so we have to check
      file_name = att$name
      writeBin(base64decode(att$contentBytes, "raw") ,  file_name)
      
    } else {
      the_attachment = AzureGraph::call_graph_endpoint(token,
                                                      paste("me/messages/", 
                                                            curr_msg$id, 
                                                            "/attachments/",
                                                            att$id,
                                                            "/$value",
                                                            sep=""))
      
      file_name = str_replace_all(att$name, "[^[:alnum:]]", " ")
      file_name = paste(file_name, ".eml", sep = "")
      write(gsub("\r\n", "\n", the_attachment), file_name)
    }
    attachment_names_and_types[[count]] = c(file_name, att$contentType)
    count = count + 1
  }
  attachment_names_and_types
}

#upload_file
upload_file = function(local_file_name, content_type, folder_id, drive_id) {
  file_content = read_file_raw(file = local_file_name)
  call_graph_url(me$token,paste ("https://graph.microsoft.com/beta/drives/",
                                 drive_id,
                                 "/items/",
                                 folder_id,
                                 ":/",
                                 gsub(" ", "%20", local_file_name),
                                 ":/content",
                                 sep=""),
                 add_headers  ("Content-Type"=content_type),
                 body=file_content, 
                 http_verb=c ("PUT"),encode = "raw")
}

#create_subfolder
create_subfolder = function (subfolder_name, folder_id, drive_id) {
  body_prep = list(name = subfolder_name,
                   folder = list(),
                   '@microsoft.graph.conflictBehavior' = "rename")
  body_json = toJSON (body_prep, auto_unbox=TRUE)
  body_json = gsub("\\[\\]", "\\{\\}", body_json)
  new_folder = call_graph_url(me$token,
                              paste ("https://graph.microsoft.com/beta/drives/",
                                     drive_id,
                                     "/items/",
                                     folder_id,
                                     "/children",
                                     sep=""),
                 add_headers  ("Content-Type"="application/json"),
                 body=body_json, 
                 http_verb=c ("POST"),encode = "raw")
  
}

#detect_tx_mail
# todo
# shall detect if a mail is a tx mail to be processed
# criteria
#  1 contains TX id (see detect_tx)
#  2 not a comment mail from planner
#  3 not the automated mail from power automate requesting security WIN feedback
detect_tx_mail = function(curr_msg) {
  nchar(detect_tx(curr_msg$subject)) > 0
}

#detect_tx
# find if the subject contains a Task ID 
# pattern is TI and then 6 digits
detect_tx = function(analyzed.text) {
  pos = regexpr(TX_MARKER, analyzed.text)
  if (pos > 0) {
    regmatches(analyzed.text,pos)
  } else {
    ""
  }
}

#team_for_tx
team_for_tx = function (tx_id) {
  temp = plans_and_groups %>% 
    filter(Marker == substring(tx_id, 3,3)) %>%
    select(Group)
  temp[[1]]
}

#plan_for_tx
plan_for_tx = function (tx_id) {
  print(paste("seeking plan for marker", substring(tx_id, 3,3)))
  temp = plans_and_groups %>% 
    filter(Marker == substring(tx_id, 3,3)) %>%
    select(Plan)
  print(paste("Plan for marker", tx_id, "is:", temp[[1]]))
  temp[[1]]
}

#retrieve_plan_id
#todo: cache this information
retrieve_plan_id = function (plan_and_group) {
  plan_name = plan_and_group[[1]]
  group_name = plan_and_group[[2]]
  group_id = my_teams$value[unlist(map(my_teams$value, function(x) x$displayName == group_name))][[1]]$id
  if(is.na(plan_name)) {
    plan_id = "NA"    
  } else {
    plan_data = call_graph_url(me$token,paste("https://graph.microsoft.com/v1.0/groups/",group_id,"/planner/plans",sep=""))
    plan_id = plan_data$value[unlist(map(plan_data$value, function(x) x$title == plan_name))][[1]]$id
  }
  group_drive = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/",group_id,"/drive", sep=""))
  group_drive_id = group_drive$id
  c(plan_id, group_id, group_drive_id)
}

det_effort = function (category_list) {
  #eliminate categories 1 and 6
  if (length(category_list[category_list %in% 'category2'] ) == 1) {
    'category2'
  } else {
    if (length(category_list[category_list %in% 'category3'] ) == 1) {
      'category3'
    } else {
      'NA'
    }
  }
}

det_problem = function (category_list) {
  if (length(category_list[category_list %in% 'category1'] ) == 1) {
    'category1'
  } else {
    'NA'
  }  
}

det_waiting = function (category_list) {
  if (length(category_list[category_list %in% 'category6'] ) == 1) {
    'category6'
  } else {
    'NA'
  } 
}

cut_date_time = function (date_time_string) {
  if (nchar(date_time_string) >= 10) {
    temp_date_time = substring(date_time_string, 1, 10)
  } else {
    temp_date_time = date_time_string
  }    
  #  format(strptime(temp_date_time, format="%Y-%m-%d"), format="%d/%m/%Y")
  strptime(temp_date_time, format="%Y-%m-%d")
}

ensure_value = function (uncertain_val) {
  if (is.null(uncertain_val)) {
    'na'
  } else {
    uncertain_val 
  }
}


#det_new_message
det_new_message = function (category_list) {
  if (length(category_list[category_list %in% 'category5'] ) == 1) {
    'category5'
  } else {
    'NA'
  }  
}

#cache of the task tables; avoids that the same planner has to be read multiple times
task_tbl_cache = dict()

# function that reads all tasks for a plan (identified by its id)
#  and extracts selected columns
extract_task_tbl = function(plan_name_and_id) {
  plan_name = plan_name_and_id[1][[1]]
  plan_id = plan_name_and_id[2][[1]]
  if (task_tbl_cache$has(plan_name)) {
    #take from the cache
    print(paste("tasks retreived from cache for plan", plan_name))
    task_tbl_cache$get(plan_name)
  } else {
    inv_tasks = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/planner/plans/",plan_id,"/tasks", sep=""))
    num_rows = length(inv_tasks$value)
    #extract assignments (that are nested)
    inv_tasks_assignments = lapply(inv_tasks$value, function(x) x$assignments)
    inv_tasks_assignments_ids = lapply(inv_tasks_assignments, names)
    inv_tasks_assignments_ids = lapply(inv_tasks_assignments_ids, function (xx) if (length(xx) == 0) "NA" else xx[[1]])  
    inv_tasks_categories = lapply(inv_tasks$value, function(x) names(x$appliedCategories))
    inv_tasks_effort = lapply(inv_tasks_categories, function(x) det_effort(x))
    inv_tasks_problem = lapply(inv_tasks_categories, function(x) det_problem(x))
    inv_tasks_waiting = lapply(inv_tasks_categories, function(x) det_waiting(x))
    inv_tasks_new_message = lapply(inv_tasks_categories, function(x) det_new_message(x))
    #compose the initial tibble that contains only the plan name repeated for each row
    tib_inv_tasks = as_tibble(list(plan=rep(plan_name, num_rows)))
    # compose the lists and clean the values (they are all lists and need sapply fun=paste to become char)
    tib_inv_tasks = tib_inv_tasks %>%
      mutate(title=sapply(lapply(inv_tasks$value, function(x) x$title), FUN=paste)) %>%
      mutate(bucketId = sapply(lapply(inv_tasks$value, function(x) x$bucketId), FUN=paste)) %>%
      mutate(createdDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(x$createdDateTime)), FUN=paste)) %>% 
      mutate(startDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(ensure_value(x$startDateTime))), FUN=paste)) %>%
      mutate(dueDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(ensure_value(x$dueDateTime))), FUN=paste)) %>%
      mutate(completedDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(ensure_value(x$completedDateTime))), FUN=paste)) %>%
      mutate(id=sapply(lapply(inv_tasks$value, function(x) x$id), FUN=paste)) %>%
      mutate(priority=sapply(lapply(inv_tasks$value, function(x) x$priority), FUN=paste)) %>%
      mutate(assignments = unlist(inv_tasks_assignments_ids)) %>%
      mutate(effort = sapply(inv_tasks_effort, FUN=paste)) %>%
      mutate(problem = sapply(inv_tasks_problem, FUN=paste)) %>%
      mutate(waiting = sapply(inv_tasks_waiting, FUN=paste)) %>%
      mutate(new_message = sapply(inv_tasks_new_message, FUN=paste)) %>%
      mutate(categories = sapply(inv_tasks_categories, FUN=paste)) %>%
      mutate(etag = sapply(lapply(inv_tasks$value, function(x) x$'@odata.etag'), FUN=paste))

    #fill in the cache
    task_tbl_cache$set(plan_name, tib_inv_tasks)    
    print(paste("tasks were retreived from planner and fed into cache for plan", plan_name))
    tib_inv_tasks
  }
}

efsaholidays = c("2000-01-01", "2020-04-09", "2020-04-10", 
                 "2020-04-13", "2020-05-01", "2020-05-21", 
                 "2020-05-22", "2020-06-01", "2020-06-02",
                 "2020-11-02","2020-12-24","2020-12-25",
                 "2020-12-28","2020-12-29","2020-12-30" ,
                 "2100-01-01")

create.calendar("EFSA", efsaholidays, weekdays=c("saturday", "sunday")) 

extract_buckets = function (plan_name_and_id) {
  plan_name = plan_name_and_id[1]
  plan_id = plan_name_and_id[2]
  buckets = call_graph_url(me$token,paste("https://graph.microsoft.com/v1.0/planner/plans/",plan_id,"/buckets", sep=""))
  num_rows = length(buckets$value)
  tib_buckets = as_tibble(list(b_plan = rep(plan_name, num_rows)))
  tib_buckets = tib_buckets %>%
    mutate(b_name=sapply(lapply(buckets$value, function(x) x$name), FUN=paste)) %>%
    mutate(b_id=sapply(lapply(buckets$value, function(x) x$id), FUN=paste))
}

extract_parameter = function (text, param_name) {
  result = ""
  if (grepl(paste(param_name, "=", sep=""),text)) {
    pattern_with_quotes = paste(param_name, "=['|\"][ |A-Z|a-z|0-9|/|\\.|-]*['|\"]", sep="")
    pattern_no_quotes = paste(param_name, "=[A-Z|a-z|0-9|/|\\.|-]+", sep="")
    hit_with_quotes = str_extract(text, pattern_with_quotes)
    hit_no_quotes = str_extract(text, pattern_no_quotes)
    if (!is.na(hit_with_quotes)) {
      result = str_remove_all(hit_with_quotes, "[\"|\']")
      text = str_remove(text, pattern_with_quotes) 
    } 
    if (!is.na(hit_no_quotes)) {
      result = hit_no_quotes
      text = str_remove(text, pattern_no_quotes) 
    }
  }
  result = str_remove(result, paste(param_name, "=", sep=""))
  c(result, text)
}

update_inv_task = function(p_id, p_etag, new_assignee) {
  print ("going to update and assign")
  print(new_assignee)
  http_body = list(assignments = list (nass = list('@odata.type' = "microsoft.graph.plannerAssignment", orderHint="N9917 U2883!")))
  body_json = toJSON (http_body, auto_unbox = TRUE)
  body_json = gsub("nass", new_assignee, body_json  )
  call_graph_url(me$token,paste ("https://graph.microsoft.com/beta/planner/tasks/",
                                 p_id,
                                 sep=""),
                 add_headers  ("If-Match"=p_etag),
                 body=body_json, 
                 http_verb=c ("PATCH"),encode = "raw")
} 

#taskify
taskify = function(curr_msg) {
  print(paste("taskify:", curr_msg$subject))
  thePlan = strsplit(curr_msg$subject, " ")[[1]][[2]]
  #use the plan to find the other information from the parameter table
  thePlanName = (plans_and_groups %>% filter(grepl(thePlan, PlanDetection, ignore.case = TRUE)) %>% select(Plan))[[1]]
  theMarker = (plans_and_groups %>% filter(grepl(thePlan, PlanDetection, ignore.case = TRUE)) %>% select(Marker))[[1]]
  
  tx_id = paste("TX", theMarker, str_pad(round(runif(1) * 1000000,0), 6, pad=0), sep="")

  raw_title = trimws(str_remove(str_remove(curr_msg$subject, "Taskify"), "TX[A-Z]"))
  
  #extract bucket if it was provided
  temp = extract_parameter(raw_title, "BUCKET")
  theBucket = temp[1]
  raw_title = temp[2]
  
  if(theBucket == "")  { #if no bucket was provided we use the deault bucket
    theBucket = (plans_and_groups %>% filter(grepl(thePlan, PlanDetection, ignore.case = TRUE)) %>% select(Defbucket))[[1]]
  }  

  #extract channel  if it was provided
  temp = extract_parameter(raw_title, "CHANNEL")
  theChannel = temp[1]
  raw_title = temp[2]
  
  if(theChannel == "")  { #if no bucket was provided we use the deault bucket
    theChannel = (plans_and_groups %>% filter(grepl(thePlan, PlanDetection, ignore.case = TRUE)) %>% select(DefChannel))[[1]]
  }  

  #extract start if it was provided
  temp = extract_parameter(raw_title, "START")
  theStartDate = temp[1]
  raw_title = temp[2]
  
  if(theStartDate != "")  { 
    temp_start_date = as.Date(theStartDate,format="%Y-%m-%d")
    if (is.na(temp_start_date)) {
      temp_start_date = as.Date(theStartDate,format="%d/%m/%Y")
    }
    if (is.na(temp_start_date)) {
      temp_start_date = as.Date(theStartDate,format="%d.%m.%Y")
    }
  } else {
    temp_start_date = Sys.Date()
  }  
  
  #extract due if it was provided
  temp = extract_parameter(raw_title, "DUE")
  theDueDate = temp[1]
  raw_title = temp[2]
  
  if(theDueDate != "")  { #if no bucket was provided we use the deault bucket
    temp_due_date = as.Date(theDueDate,format="%Y-%m-%d")
    if (is.na(temp_due_date)) {
      temp_due_date = as.Date(theDueDate,format="%d/%m/%Y")
    }
    if (is.na(temp_due_date)) {
      temp_due_date = as.Date(theDueDate,format="%d.%m.%Y")
    }
  } else {
    temp_due_date = offset(Sys.Date(), DEFAULT_OFFSET, "EFSA")
  }

  #extract sharepoint folder path  if it was provided
  temp = extract_parameter(raw_title, "PATH")
  sp_folder_path = temp[1]
  raw_title = temp[2]
  
  if(sp_folder_path == "")  { #if no bucket was provided we use the deault bucket
    sp_folder_path = (plans_and_groups %>% filter(grepl(thePlan, PlanDetection, ignore.case = TRUE)) %>% select(FolderPath))[[1]]
  }  
  
  #extract assignee  if it was provided
  temp = extract_parameter(raw_title, "ASSIGN")
  new_assignee = temp[1]
  raw_title = temp[2]
  if (new_assignee !="") {
    new_assignee = unlist(user_tib %>% filter(user_names == new_assignee) %>% select(user_ids))
  }
  
  #steps: first extract parameter, then check if parameter fits one of the allowed names, then call (copy it from AutoAssign) update_inv_task
  theTitle = paste("(", tx_id, ") ", 
                   trimws(substr(raw_title, 1, 80)),
                   sep="")

#todo: gsub to clean filename
  theFilename = trimws(substr(theTitle, 1, 40))
  
  theShortFilename = trimws(substr(theTitle, 1, 30))
      
  all_ids = retrieve_plan_id(c(plan_for_tx(tx_id), team_for_tx(tx_id)))

  ##task creation
  planId = all_ids[1]
  if(!is.na(theBucket)) {
    bucketTbl = extract_buckets(c(thePlanName, planId))
    bucketId = (bucketTbl %>% filter(b_name == theBucket) %>% select(b_id))[[1]]
    # create new task
    new_task_obj = list (planId = planId,
                         bucketId = bucketId,
                         title = theTitle,
                         startDateTime = format(temp_start_date, format="%Y-%m-%dT%H:%M:%SZ"),
                         dueDateTime = format(temp_due_date, format="%Y-%m-%dT%H:%M:%SZ"))
    if (new_assignee != "") {
      new_task_obj = c(new_task_obj, list(assignments = list (nass = list('@odata.type' = "microsoft.graph.plannerAssignment", orderHint="N9917 U2883!"))))
    } 
  
    new_task_json = toJSON (new_task_obj, auto_unbox = TRUE)
    new_task_json = gsub("nass", new_assignee, new_task_json  )
    
    print(new_task_json)
    new_task = call_graph_url(me$token,"https://graph.microsoft.com/beta/planner/tasks",  #we use beta because it can handle the priority field
                              body=new_task_json, http_verb=c ("POST"),encode = "raw")  
    
  }
  
  drive_id = all_ids[3]  

  sp_folder_path = gsub(" ", "%20", sp_folder_path)
    # create folder for the invoice checklist
  parent_folder = call_graph_url(me$token,
                                 paste("https://graph.microsoft.com/v1.0/drives/",
                                       drive_id,
                                       "/root:",
                                       sp_folder_path,
                                       sep=""))
                                 
  
  new_folder_obj = list (name  = theShortFilename,
                         folder = list(),
                         '@microsoft.graph.conflictBehavior' = "rename")
  new_folder_json = toJSON (new_folder_obj, auto_unbox = TRUE)
  new_folder_json = str_replace(new_folder_json, fixed("[]"), "{}")
#todo: pay_folder are fixed strings in invoice importer; here they have to be found using info from plans_and_groups
  new_folder = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/drives/",drive_id,"/items/",parent_folder$id,"/children", sep=""),  
                              body=new_folder_json, http_verb=c ("POST"),encode = "raw")  
  
  # save the mail to the new folder
  print(paste("Saving the taskify mail to SP with subject ", curr_msg$subject, sep=""))
  # save mail to hd
  write_mail(curr_msg, tx_id = tx_id)
  # upload message to sp
  upload_file(generate_file_name(curr_msg, tx_id = tx_id), "message/rfc822", new_folder$id, drive_id)      
  # if there are attachments: save attachments to hd
  #                           create subfolder for attachments
  #                           upload attachments
  if (curr_msg$hasAttachments) {
    attachment_names_and_types = save_attachments(curr_msg)
    new_folder_att = create_subfolder(generate_folder_name(curr_msg), new_folder$id, drive_id)
    sapply(attachment_names_and_types, function (x) upload_file(x[[1]],x[[2]], new_folder_att$id, drive_id))
    delete_temp_attachment_files(attachment_names_and_types)
  }
  #delete the local copy
  delete_temp_mail_file(curr_msg, tx_id = tx_id)

  #prepare the links
  folder_url_enc = str_replace_all(new_folder$webUrl, fixed(":"), "%3A")
  folder_url_enc = str_replace_all(folder_url_enc, fixed("."), "%2E")
  
  
  ##insert conversation post
  group_id = all_ids[2]
  channels = call_graph_url(me$token,
                            paste("https://graph.microsoft.com/v1.0/teams/",
                                  group_id,
                                  "/channels",
                                  sep=""))
  channel_names = sapply(channels$value, function(x) x$displayName)
  def_channel_pos = which(sapply(channel_names, function(x) grepl(theChannel, x, ignore.case = TRUE)))[[1]]
  channel_id = channels$value[[def_channel_pos]]$id
  
  unique_body = AzureGraph::call_graph_endpoint(token,paste("me/mailFolders/inbox/messages/",curr_msg$id,"?$select=uniqueBody",sep=""))
  
  if (!is.na(theBucket)) {
    new_conversation_prep = list(subject = theTitle,
                                 body=list(contentType = "html", 
                                           content=paste("A new task ",
                                                         "(<a href='tempTaskUrlUnenc'>link</a>)",
                                                         " was created from the mail stored under the following folder (<a href='tempWebUrlUnenc'>link</a>).<br>",
                                                         unique_body$uniqueBody$content,
                                                         sep="")))
  } else {
    new_conversation_prep = list(subject = theTitle,
                                 body=list(contentType = "html", 
                                           content=paste("The following request was created from a mail stored under the following folder (<a href='tempWebUrlUnenc'>link</a>).<br>",
                                                         unique_body$uniqueBody$content,
                                                         sep="")))
  }
  
  new_conversation_json = toJSON(new_conversation_prep,auto_unbox=TRUE)
  new_conversation_json = str_replace_all(new_conversation_json , "tempWebUrlUnenc", new_folder$webUrl)
  new_conversation_json = str_replace_all(new_conversation_json , "tempWebUrlEnc", folder_url_enc)
  
  if(!is.na(theBucket)) {
    new_conversation_json = str_replace_all(new_conversation_json , "tempTaskUrlUnenc", paste("https://tasks.office.com/EFSA815.onmicrosoft.com/en-gb/Home/Task/", new_task$id,sep=""))
  }
  
  new_conversation = call_graph_url(me$token,
                                    paste("https://graph.microsoft.com/beta/teams/",
                                          group_id,
                                          "/channels/",
                                          channel_id,
                                          "/messages", 
                                          sep=""),  
                                    body=new_conversation_json, 
                                    http_verb=c ("POST"),
                                    encode = "raw")  
  
  ##update the task detilas to insert the link 
  if (!is.na(theBucket)) {
    #prepare the description
    desc_prep = list()
    #the key of the attachment is the encoded URL
    #it does not work with the original URL and again we have to
    #substitute after toJSON
    body_prep = list (previewType = "noPreview",
                      description="See attached mail for request details",
                      references = list(tempWebUrlEnc = list(alias = "Checklist folder",
                                                             '@odata.type' = "microsoft.graph.plannerExternalReference",
                                                             type = "Other")))
    body_json = toJSON (body_prep, auto_unbox=TRUE)
    body_json = str_replace_all(body_json , "tempWebUrlUnenc", new_folder$webUrl)
    body_json = str_replace_all(body_json , "tempWebUrlEnc", folder_url_enc)
    #body_json = paste("#JSON",body_json)
    print(body_json)
    
    #get details for etag
    new_details = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/planner/tasks/",new_task$id,"/details",sep=""))
    etag = new_details$'@odata.etag'
    call_graph_url(me$token,paste ("https://graph.microsoft.com/beta/planner/tasks/",
                                   new_task$id,
                                   "/details",sep=""),
                   add_headers  ("If-Match"=etag),
                   body=body_json, 
                   http_verb=c ("PATCH"),encode = "raw")
  }

  
}


#update task (sets new message category, eliminates waiting)
update_inv_task = function(task_id) {
  planner_task = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/planner/tasks/",task_id,sep=""))
  etag = planner_task$'@odata.etag'
  
  previousCategories = planner_task$appliedCategories
  previousCategories[["category5"]] = TRUE
  if (CANCEL_WAITING_UPON_NEW_MAIL) {
    previousCategories[["category6"]] = FALSE 
  }
  new_categories = list(appliedCategories = previousCategories)

  #get details for etag
  call_graph_url(me$token,paste ("https://graph.microsoft.com/beta/planner/tasks/",
                                 task_id,
                                 sep=""),
                 add_headers  ("If-Match"=etag),
                 body=toJSON (new_categories, auto_unbox = TRUE), 
                 http_verb=c ("PATCH"),encode = "raw")
  

}

#set_read_flag
set_read_flag = function (curr_msg) {
  print(paste("Setting read flag for mail with subject ", curr_msg$subject, sep=""))
  #set the read flag to true
  body_prep = list(isRead = TRUE)
  
  call_graph_url(me$token,
                 paste("https://graph.microsoft.com/beta/me/messages/", 
                       curr_msg$id, sep=""),
                 body=toJSON (body_prep, auto_unbox=TRUE), 
                 add_headers  ("Authorization"=token$credentials$access_token, "Content-Type"="application/json"),
                 http_verb=c ("PATCH"),encode = "raw")
  
  #store in a global variable that read flag was set; see the main function for the explanation of the need of this
  #the information will be used at the end of elaborate_mail
  read_flag_was_set <<- TRUE
}

#detect_current_user_assignee
# detects the azure user id of the current user (a 30+ alphanumeric code)
# the planner task contains the azure user id of the assignee
# and we want to compare it to the azure user id of the current user
detect_current_user_assignee = function () {
  curr_user = Sys.getenv("USERNAME")
  if (curr_user %in% sapply(user_assignments_names, function(x) x[[1]])) {
    short_name = user_assignments_names[which(sapply(user_assignments_names, function(x) x[1] == curr_user))][[1]][2]
    o365_user_code = user_tib %>% 
      filter(user_names == short_name) %>%
      select(user_ids)
    o365_user_code[[1]]
  } else {
    "NA"    
  }
}

#elaborate_tx
# function that elaborates mails that have a transaction id
# tx ids are detected by detect_tx
elaborate_tx = function(curr_msg) {
  print(paste("Elaborate TX for mail with subject ", curr_msg$subject, sep=""))
  tx_id = detect_tx(curr_msg$subject)
  if (tx_id != "") {
    all_ids = retrieve_plan_id(c(plan_for_tx(tx_id), team_for_tx(tx_id)))
    if (!is.na(plan_for_tx(tx_id))) {
      #search for task and check if it is assigned to the current user
      plan_id = all_ids[1]
      task_tbl = extract_task_tbl(c(plan_for_tx(tx_id), plan_id))
      right_task = task_tbl %>% filter(grepl(tx_id, title, ignore.case = TRUE))
      task_id = right_task$id
    } else {
      task_id = "NO PLAN"
    }
    if (length(task_id) !=0) { #protection against IDs that have been deleted in the planner task 
      if ((!DONT_MOVE_TO_SP_OTHERS ||
           stri_cmp_eq(right_task$assignments , "NA") ||
           stri_cmp_eq(right_task$assignments,detect_current_user_assignee())) ||
          stri_cmp_eq(task_id , "NO PLAN")) {
        #search for target folder; 
        drive_id = search_tx_drive(tx_id, team_for_tx(tx_id))
        folder_id = search_tx_folder(tx_id, drive_id)
        #search for the mail in the folder by ID
        saved_mail = search_mail_in_folder(curr_msg, folder_id, drive_id)    
        #verify if mail was already saved in sp
        if (is.null(saved_mail)) {
          print(paste("Transferring to SP mail with subject ", curr_msg$subject, sep=""))
          # save mail to hd
          write_mail(curr_msg)
          # upload message to sp
          new_file_id = upload_file(generate_file_name(curr_msg), "message/rfc822", folder_id, drive_id)      
          # if there are attachments: save attachments to hd
          #                           create subfolder for attachments
          #                           upload attachments
          if (curr_msg$hasAttachments) {
            attachment_names_and_types = save_attachments(curr_msg)
            new_folder = create_subfolder(generate_folder_name(curr_msg), folder_id, drive_id)
            sapply(attachment_names_and_types, function (x) upload_file(x[[1]],x[[2]], new_folder$id, drive_id))
            delete_temp_attachment_files(attachment_names_and_types)
          }
          #delete the local copy
          delete_temp_mail_file(curr_msg)
          # flag the planner task, but only if the message was not sent by the team
          if(!detect_own_message(curr_msg)) {
            update_inv_task(task_id)
          }
          # search conversation with this id
          # add post to conversation with link to mail
          # adding posts was blocked after I received negative feedback about this feature
          to_teams = grepl("TOTEAMS", curr_msg$subject, ignore.case = TRUE)
          if (to_teams) {
            theSubject = str_remove(str_remove(curr_msg$subject, "TOTEAMS"), "toteams")
            group_id = all_ids[2]
            theChannel = (plans_and_groups %>% filter(Marker == substr(tx_id, 3, 3)) %>% select(DefChannel))[[1]]
            print(paste("going to insert new reply in channel", theChannel))
            channels = call_graph_url(me$token,
                                      paste("https://graph.microsoft.com/v1.0/teams/",
                                            group_id,
                                            "/channels",
                                            sep=""))
            channel_names = sapply(channels$value, function(x) x$displayName)
            right_channel = sapply(channel_names, function(x) grepl(theChannel, x, ignore.case = TRUE))
            if(TRUE %in% right_channel) {
              print(paste("the right channel was found"))
              def_channel_pos = which(right_channel)[[1]]
              channel_id = channels$value[[def_channel_pos]]$id
              
              messages = call_graph_url(me$token,
                                        paste("https://graph.microsoft.com/beta/teams/",
                                              group_id,
                                              "/channels/",
                                              channel_id,
                                              "/messages?$top=50", 
                                              sep=""))
              subjects = sapply(messages$value, function(x) x$subject)
              
              print(paste("before loop: length of subject list of new channel message slice", length(messages$value)))
              print(paste("before loop: length of subject list of channel messages", length(subjects)))
              
              skip_value = 0
              subject_pos = which(grepl(tx_id, subjects, ignore.case = TRUE))
              
              #todo: check regularly if MS implemented filter on the 
              #query for channel messages; if that is the case, eliminate
              #the loop and filter directly by 'contains the tx_id' 
              while (length(messages$value) > 0 & length(subject_pos) == 0) {
                skip_value = skip_value + 50
                messages = call_graph_url(me$token,
                                          paste("https://graph.microsoft.com/beta/teams/",
                                                group_id,
                                                "/channels/",
                                                channel_id,
                                                "/messages?$top=50&$skip=", 
                                                skip_value,
                                                sep=""))
                subjects = sapply(messages$value, function(x) x$subject)
                print(paste("length of subject list of new channel message slice", length(messages$value)))
                print(paste("length of subject list of channel messages", length(subjects)))
                subject_pos = which(grepl(tx_id, subjects, ignore.case = TRUE))
              }
              
              if(length(subject_pos) > 0) {
                print(paste("number of messages found", length(subject_pos)))
                message_id = messages$value[[subject_pos[[1]]]]$id #the [[1]] is needed because there might be more than one message with the code; we use the first one
                unique_body = AzureGraph::call_graph_endpoint(token,paste("me/mailFolders/inbox/messages/",curr_msg$id,"?$select=uniqueBody",sep=""))
                
                new_conversation_prep = list(body=list(contentType = "html", 
                                                       content=paste("New mail from ",
                                                                     curr_msg$from$emailAddress$name, 
                                                                     " sent on ",
                                                                     curr_msg$sentDateTime,
                                                                     " (<a href='tempWebUrlUnenc'>link</a>)<br>",
                                                                     theSubject,
                                                                     "<br>",
                                                                     unique_body$uniqueBody$content,
                                                                     sep="")))
                
                new_conversation_json = toJSON(new_conversation_prep,auto_unbox=TRUE)
                new_conversation_json = str_replace_all(new_conversation_json , "tempWebUrlUnenc", new_file_id$webUrl)
                
                new_conversation = call_graph_url(me$token,
                                                  paste("https://graph.microsoft.com/beta/teams/",
                                                        group_id,
                                                        "/channels/",
                                                        channel_id,
                                                        "/messages/",
                                                        message_id,
                                                        "/replies",
                                                        sep=""),  
                                                  body=new_conversation_json, 
                                                  http_verb=c ("POST"),
                                                  encode = "raw")  
                
              }
            }
            
          }
        }
      }
    }
    #flag mail "read"; this is done outside the if because it shall be 
    # flagged in the inbox in case the user is the second person to 
    # find this mail; must-read mails (e.g. VIP) are never flagged
    if(FLAG_TX_READ && !must_read(curr_msg)) {
      set_read_flag(curr_msg)  
    }
  }
}



##test
if (mode == "test_TX") {
}

if (mode == "test_extract_parameter") {
  print(extract_parameter("Taskify TXC BUCKET='to do' Do it", "BUCKET"))
}

if (mode == "test_search_tx") {
  drive_id = search_tx_drive("TI052527", "Tech Specific Contract")
  folder_id = search_tx_folder("TI052527",drive_id)
}

if (mode == "test_mime") {
  write_mail(curr_msg)
}

if (mode == "test_search_in_folder") {
  print(search_mail_in_folder (curr_msg, "01RTSAYT4J5ICNVIZSMZFKWESBYQABFUT5","b!wt8aOM7MVkS0m_DhdwDa3mzwuQ5ZoQdDjsFgvnewpev2F8IiPURjSriJ-blxYH7a") )
}

if (mode == "test_save_att") {
  save_attachments(curr_msg)
}


if (mode == "test_set_read") {
}

if (mode == "test_upload_mail") {
  upload_file("20200413 Re  Your reply to all is needed for submitted contract request  Test Contract  TI052527 MIDAAMkAD.eml","01RTSAYT4J5ICNVIZSMZFKWESBYQABFUT5","b!wt8aOM7MVkS0m_DhdwDa3mzwuQ5ZoQdDjsFgvnewpev2F8IiPURjSriJ-blxYH7a")
}

if (mode == "test_create_folder") {
  create_subfolder ("pippo","01RTSAYT4J5ICNVIZSMZFKWESBYQABFUT5","b!wt8aOM7MVkS0m_DhdwDa3mzwuQ5ZoQdDjsFgvnewpev2F8IiPURjSriJ-blxYH7a") 
}

if (mode == "test_extract_tasks") {
  plan_id = retrieve_plan_id(c(plan_for_tx("TXC123456"), team_for_tx("TXC123456")))[1]
  task_tbl = extract_task_tbl(c(plan_for_tx("TSC123456"), plan_id))
}


if (mode == "test_update_task") {
  update_inv_task("MHXP4_DpZUWCrgzrtucKvZYAFQhb", c())
}

##main
if (mode == "prod") {
  main()
}

if (mode == "onemail") {
  main_step_by_step()
}
