# utility that retrieves planner tasks for the TSTA team
# computes effort points
# and saves them to two CSV:
# all tasks contains all tasks
# active tasks contains tasks that are not completed, nor waiting nor problems
# csv can be downloaded using the file/open menu above

#configuration parameters
#gather_descriptions indicates if the descriptions of the active tasks shall be downloaded
#in 15% of cases gathering descriptions leads to an error message

load(paste("/home/docker/myEnvironmentMailReader.RData",sep=""))

print(paste("*** Mail reader starting at: ", Sys.time()))

gather_descriptions = FALSE

##configuration: to be changed by every user

#execution mode of the script (prod or onemail)
mode = "prod" # mode="prod" requests to elaborate all unread mails; mode="onemail" asks to elaborate only the first unread mail

#activate/deactivate filtering functions
SET_READ_AUTOMATED_MESSAGES = TRUE #set read messages from automated senders (list see below)
SET_READ_BY_SUBJECT = TRUE #set read mails with special subjects specified below
SET_READ_INVITATIONS = TRUE #set read invitations and replies to invitations
SET_READ_PLANNER = TRUE #set read planner mails
SET_READ_SERVICE_NOW_ALL = FALSE #set read all service now mails
SET_READ_SERVICE_NOW_EXCEPT_HOLD = TRUE #todo: allow to read service now on-hold and signed/closed
SET_READ_ALL_CC = TRUE #set read all mails that were not sent to the current user and some relevant TS TA groups
ELABORATE_TX_MAILS = TRUE #elaborate mails with a transaction code (move mail to folder, set new-message flag on task, eliminate waiting flag)
FLAG_TX_READ = TRUE #if a mail pertaining a transaction is found, it is set read
DONT_MOVE_TO_SP_OTHERS = FALSE #move to sharepoint only mails for tasks that are assigned to current user or unassigned; avoids that colleagues are suprised by mails in their folder
CANCEL_WAITING_UPON_NEW_MAIL = TRUE #if new mail received, eliminate "Waiting" flag

DEFAULT_OFFSET = 7 #days from now that the due date is set to if no due date is indicated in taskify

# list of people treated like VIP; for their mails readFlag is never set
VIP_senders = "selomey.yamadjako@efsa.europa.eu;paul.devalier@efsa.europa.eu;chiara.bianchi@efsa.europa.eu;fabrizio.abbinante@efsa.europa.eu;sosanna.tasiou@efsa.europa.eu;giovanni.fuga@efsa.europa.eu" #todo complete this list

#senders of automated mails (readFlag is set)
automated_senders = "noreply@researchcircle-gartner.com;suggestions@clients.gartner.com;no-reply@sharepointonline.com;efsa.jiveon.com; noreply@Planner.Office365.com; ec-fp-internet-services-do-not-reply@ec.europa.eu; noreply@efsa.europa.eu; noreply@yammer.com; noreply@sciforma.net; u4u-news-com-request@lists.u4unity.eu; REP-PERS-OSP-U4U@ec.europa.eu; noreply@email.teams.microsoft.com; azuredevops@microsoft.com"

#subject components for which readFlag is set
exclude_subjects_beginning_with = c("Information Security Monthly Management Update",
                                    "Automatic reply:",
                                    "Re: Comments on task",
                                    "ABAC invoice report for",
                                    "ABAC commitment report for",
                                    "Request to TS Assurance created",
                                    "Comments on task",
                                    "Your reply-to-all is needed for submitted contract request",
                                    "Taskify")

exclude_subjects_containing = c()

exclude_subjects_ending_with = c()

#groups the current user belongs to (mails to these addresses and to the current user are considered not to be cc)
to_groups = c("TeamTSTAServices@EFSA815.onmicrosoft.com",
              "_AllStaff@efsa.europa.eu",
              "DTS@group.efsa.europa.eu")

#ts ta team members mails; mails from these senders do not flag the task as "new message"
tsta_member_mails = c("alessandro.gazzetta@efsa.europa.eu",
                      "mavra.nikolopoulou@efsa.europa.eu",
                      "despoina.papadopoulou@efsa.europa.eu")


source(paste("/home/docker/MailReaderCreator.R",sep=""))

save.image(file=paste("/home/docker/myEnvironmentMailReader.RData",sep=""))

print("image saved, job completed")

print(paste("*** Mail reader closing at: ", Sys.time()))

