;;; org-outlook.el --- sync events from outlook calendar to org -*- lexical-binding: t -*-

;; Copyright (C) 2022 Ian FitzPatrick

;; Author: Ian FitzPatrick ian@ianfitzpatrick.eu
;; URL: github.com/ifitzpat/org-outlook
;; Version: 0.1.0
;; Package-Requires: ((emacs "27.1") (org-ml "6.0.2") (html2org "0.1") (request "0.3.3") (org-msg "4.0") (web-server "0.1.2"))
;; Keywords: calendar outlook org-mode

;; This file is not part of GNU Emacs

;; This file is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation; either version 3, or (at your option)
;; any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; For a full copy of the GNU General Public License
;; see <https://www.gnu.org/licenses/>.

;;; Commentary:
;;
;; [FIXME package description]

;;; Code:

(require 'org-ml)
(require 'html2org)
(require 'request)
(require 'org-msg)
(require 'plstore)
(require 'web-server)

(defconst org-outlook-resource-url "https://graph.microsoft.com/Calendars.ReadWrite")
(defconst org-outlook-events-url "https://graph.microsoft.com/v1.0/me/calendarview")
(defconst org-outlook-events-create-url "https://graph.microsoft.com/v1.0/me/calendar/events")

(defvar org-outlook-local-timezone "Europe/Berlin" "Your timezone")
(defvar org-outlook-token-cache-file "~/.cache/outlook.plist" "Path to a plist file to keep the encrypted secret tokens")
(defvar org-outlook-sync-start 14 "How many days 'in the past' should be synced?")
(defvar org-outlook-sync-end 90 "How many days 'in the future' should be synced?")
(defvar org-outlook-file "~/.emacs.d/outlook.org")

(defvar org-outlook-client-id "3df0b076-dc9c-48f8-b940-a271ed0bb14b" "Microsoft Entra App Registration Client ID. You can use the default or provide your own if you prefer (see README.org for details)")
(defvar org-outlook-tenant-id "organizations" "If you provide your own App Registration you can optionally set this to only your outlook tenant (see README.org for details).")

(defvar org-outlook-auth-url (format "https://login.microsoftonline.com/%s/oauth2/v2.0/authorize" org-outlook-tenant-id))
(setq org-outlook-token-url (format "https://login.microsoftonline.com/%s/oauth2/v2.0/token" org-outlook-tenant-id))

(defun org-outlook-generate-random-string (length)
  "Generate a random string of LENGTH consisting of URL-safe characters."
  (let ((chars "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~")
        (result ""))
    (dotimes (_ length)
      (setq result (concat result (string (aref chars (random (length chars)))))))
    result))

(setq org-outlook-code-verifier (org-outlook-generate-random-string 43))
(setq org-outlook-code-challenge (base64url-encode-string (secure-hash 'sha256 org-outlook-code-verifier nil nil t)))
;(setq org-outlook-code-challenge org-outlook-code-verifier)

;; OAuth state tracking variables
(defvar org-outlook--auth-server nil "Web server instance for OAuth.")
(defvar org-outlook--auth-complete nil "Semaphore for OAuth completion.")
(defvar org-outlook--auth-timer nil "Timer for OAuth timeout.")

;;
(defun n-days-ago (&optional n)
  (let* ((days (or n 90))
         (timestamp (time-subtract (current-time) (days-to-time days))))
    (list
     (string-to-number (format-time-string "%m" timestamp))
     (string-to-number (format-time-string "%d" timestamp))
     (string-to-number (format-time-string "%Y" timestamp)))))

(defun token-timed-out (&optional token)
  (let* ((token (or token "access"))
	 (org-outlook-token-cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
	 (auth-timestamp (plist-get (cdr (plstore-get org-outlook-token-cache token)) :timestamp)))
    (plstore-close org-outlook-token-cache)
    (if auth-timestamp
	(time-less-p (time-add (parse-iso8601-time-string auth-timestamp) (seconds-to-time 3599))  (current-time))
      nil)))

(defun token-cache-exists ()
  (file-regular-p org-outlook-token-cache-file))

(defun org-outlook-start ()
  (format-time-string "%Y-%m-%d" (time-subtract (current-time)(days-to-time org-outlook-sync-start))))

(defun org-outlook-end ()
  (format-time-string "%Y-%m-%d" (time-subtract (current-time)(days-to-time (* -1 org-outlook-sync-end)))))

(defun org-outlook-start-time-from-timestamp (ts)
  (let* ((tobject (cadr (org-timestamp-from-string ts)))
	 (year-start (plist-get tobject :year-start))
	 (month-start (plist-get tobject :month-start))
	 (day-start (plist-get tobject :day-start))
	 (hour-start (plist-get tobject :hour-start))
	 (minute-start (plist-get tobject :minute-start)))
    (format "%s-%02d-%02dT%02d:%02d:00" year-start month-start day-start hour-start minute-start)))

(defun org-outlook-end-time-from-timestamp (ts)
  (let* ((tobject (cadr (org-timestamp-from-string ts)))
	 (year-end (plist-get tobject :year-end))
	 (month-end (plist-get tobject :month-end))
	 (day-end (plist-get tobject :day-end))
	 (hour-end (plist-get tobject :hour-end))
	 (minute-end (plist-get tobject :minute-end)))
    (format "%s-%02d-%02dT%02d:%02d:00" year-end month-end day-end hour-end minute-end)))


(defun process-token-callback (request)
  "Process OAuth callback and signal completion."
  (with-slots (process headers) request
    (ws-response-header process 200 '("Content-type" . "text/html"))
    (let ((auth-code (cdr (assoc "code" (headers request)))))
      (if auth-code
          (progn
            (org-outlook-set-token-field "auth" nil `(:auth ,auth-code))
            (setq org-outlook--auth-complete t)
            (when org-outlook--auth-timer
              (cancel-timer org-outlook--auth-timer))
            (process-send-string process
                                 "<html><body><h2>Authentication successful!</h2><p>You can close this window and return to Emacs.</p></body></html>"))
        (process-send-string process
                             "<html><body><h2>Authentication failed</h2><p>No authorization code received.</p></body></html>")))))

(defun start-auth-code-server ()
  (ws-start
   '(((:GET . ".*") .
      (lambda (request)
	(process-token-callback request)
	)))
   9004))

(defun org-outlook-set-token-field (type public secret)
  (let
      ((org-outlook-token-cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
       (plstore-encrypt-to org-outlook-gpg-recipient))
    (plstore-put org-outlook-token-cache type public secret)
    (plstore-save org-outlook-token-cache)
    (plstore-close org-outlook-token-cache)
    ))

(defun org-outlook-request-authorization ()
  "Request OAuth authorization without blocking Emacs.
Opens the browser and waits for the callback with a 5-minute timeout.
Returns the authorization code on success."
  (let* ((outlook-auth-url
          (concat org-outlook-auth-url
                  "?client_id=" (url-hexify-string org-outlook-client-id)
                  "&response_type=code"
		  "&code_challenge=" org-outlook-code-challenge
		  "&code_challenge_method=S256"
                  "&redirect_uri=" (url-hexify-string "http://localhost:9004")
                  "&scope=" (url-hexify-string (concat "offline_access " org-outlook-resource-url)))))

    ;; Reset state
    (setq org-outlook--auth-complete nil)

    ;; Start server
    (setq org-outlook--auth-server (start-auth-code-server))

    ;; Set timeout (5 minutes)
    (setq org-outlook--auth-timer
          (run-at-time 300 nil
                       (lambda ()
                         (when (not org-outlook--auth-complete)
                           (when org-outlook--auth-server
                             (ws-stop org-outlook--auth-server))
                           (setq org-outlook--auth-server nil)
                           (error "OAuth authorization timed out after 5 minutes")))))

    ;; Open browser (non-blocking)
    (if (eq system-type 'gnu/linux)
	(browse-url-xdg-open outlook-auth-url)
        (browse-url outlook-auth-url))

    (message "Please complete authentication in your browser (you have 5 minutes)...")

    ;; Wait for completion with non-blocking checks
    (while (not org-outlook--auth-complete)
      (accept-process-output nil 0.1))

    ;; Cleanup
    (when org-outlook--auth-server
      (ws-stop org-outlook--auth-server))
    (setq org-outlook--auth-server nil)
    (when org-outlook--auth-timer
      (cancel-timer org-outlook--auth-timer))
    (setq org-outlook--auth-timer nil)

    (message "Authentication successful!")
    (org-outlook-auth-token)))

(defun org-outlook-auth-token ()
  (let*
      ((org-outlook-token-cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
       (token (plist-get (cdr (plstore-get org-outlook-token-cache "auth")) :auth)))
    (plstore-close org-outlook-token-cache)
    token))

(defun org-outlook-access-token ()
  (let*
      ((org-outlook-token-cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
       (token (plist-get (cdr (plstore-get org-outlook-token-cache "access")) :access)))
    (plstore-close org-outlook-token-cache)
    token))

(defun org-outlook-refresh-token ()
  (let* ((org-outlook-token-cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
         (token (plist-get (cdr (plstore-get org-outlook-token-cache "refresh")) :refresh)))
    (plstore-close org-outlook-token-cache)
    (and token (not (token-timed-out "refresh")) token)))

(defun org-outlook-request-access-token ()
  (let* ((refresh_token (org-outlook-refresh-token))
	 (auth_code (or refresh_token
			(progn
			  (org-outlook-request-authorization)
			  ;(message "retrieved auth token")
    			  ;(org-outlook-auth-token)
			  )
			)))
    ;(message auth_code)
    (request org-outlook-token-url
      :type "POST"
      :sync t
      :data `(("tenant" . ,org-outlook-tenant-id)
	      ("client_id" . ,org-outlook-client-id)
	      ("scope" . ,(concat "offline_access " org-outlook-resource-url))
	      ,(if refresh_token
		   `("refresh_token" . ,refresh_token)
		 `("code" . ,auth_code))
              ("redirect_uri" . "http://localhost:9004")
	      ,(if refresh_token
		   '("grant_type" . "refresh_token")
	         '("grant_type" . "authorization_code"))
	      ("code_verifier" . ,org-outlook-code-verifier))
      :parser 'json-read
      :success (cl-function
		(lambda (&key data &allow-other-keys)
		  (when data
		    (org-outlook-set-token-field "access" `(:timestamp ,(format-time-string "%Y-%m-%dT%H:%M:%S" (current-time))) `(:access ,(assoc-default 'access_token data)))
		    (org-outlook-set-token-field "refresh" `(:timestamp ,(format-time-string "%Y-%m-%dT%H:%M:%S" (current-time))) `(:refresh ,(assoc-default 'refresh_token data)))
		    )))
      :error (cl-function
	      (lambda (&rest args &key error-thrown &allow-other-keys)
		(message "Error getting access token: %S" error-thrown))))))


(defun org-outlook-bearer-token ()
  (if (or (token-timed-out)(not (token-cache-exists)) )
      (org-outlook-request-access-token))
  (org-outlook-access-token))

(defun org-outlook--skip (url)
  (string-match "%24skip=\\([0-9]*\\)" url)
  (string-to-number (match-string 1 url)))

(defun org-outlook-update-meeting-id (id) ;; unused?
  (interactive)
  (let ((elm (org-ml-parse-element-at (point))))
    (->> elm
	 (org-ml-update (lambda (hl)
			  (org-ml-headline-set-node-property "ID" id hl))))))

(defun org-outlook-get-appointment-property (prop)
  ;; (message "org-outlook-get-appointment-property called in buffer: %s, mode: %s"
  ;;          (buffer-name) major-mode)
  (unless (derived-mode-p 'org-mode)
    (message "Switching to org-mode in buffer: %s" (buffer-name))
    (let ((org-inhibit-startup t)) (org-mode)))
  (let ((elm (org-ml-parse-element-at (point))))
    (setq mytest elm)
    (->>
     elm
     (org-ml-headline-get-node-property prop))))


(defun org-outlook-calendar-create-or-update-event ()
  (interactive)
  (let* ((elm (org-ml-parse-element-at (point)))
	 (title (car (append (->> elm (org-ml-get-property :title)) nil))) ; TODO Fix
	 (body (org-msg-export-as-html (mapconcat #'concat (->> elm (org-ml-headline-get-contents (list :log-into-drawer org-log-into-drawer :clock-into-drawer org-clock-into-drawer :clock-out-notes org-log-note-clock-out) ) (-map #'org-ml-to-string) ) "") ))
	 (meeting-time (->> elm (org-ml-headline-get-node-property "MEETING-TIME"))) ; TODO FIX
	 (start (org-outlook-start-time-from-timestamp meeting-time))
	 (end (org-outlook-end-time-from-timestamp meeting-time))
					;(categories ["CAT1" "CAT2"])
	 (categories (vector))
	 (id (->> elm (org-ml-headline-get-node-property "OUTLOOK-ID")))
	 (method (if id "PATCH" "POST"))
	 (location (->> elm (org-ml-headline-get-node-property "LOCATION")))
	 (attendees (->> elm (org-ml-headline-get-node-property "INVITEES")))
	 (teamsmeeting t))

    (request (if id (concat org-outlook-events-create-url "/" id) org-outlook-events-create-url)
      :type method
      :data (json-encode `(("subject" . ,title)
			   ("body" . (("contentType" . "HTML")("content" . ,body)))
			   ("start" . (("dateTime" . ,start)("timeZone" . ,org-outlook-local-timezone)))
			   ("end" . (("dateTime" . ,end)("timeZone" . ,org-outlook-local-timezone)))
			   ("location" . (("displayName" . ,location)))
			   ("attendees" . ,(org-outlook-create-attendees-list attendees))
			   ,(if categories `("categories" . ,categories))
			   ("isOnlineMeeting" . ,teamsmeeting)
			   ("onlineMeetingProvider" . "teamsForBusiness")))

      :headers `(("Authorization" . ,(concat "Bearer " (org-outlook-bearer-token)))
		 ("Content-Type" . "application/json"))
      :parser 'json-read
      :error (cl-function
	      (lambda (&rest args &key error-thrown &key response &key data &allow-other-keys)
		(setq my-data data)
		(setq my-response response)
		(message "Error creating event: %S" error-thrown)))
      :success (cl-function
		(lambda (&key data &allow-other-keys)
		  (when data
		    (message "Created event")))))))

(if (listp 'org-agenda-files)
    (add-to-list 'org-agenda-files org-outlook-file)
  (setq org-agenda-files (list org-outlook-file)))

(add-to-list 'thing-at-point-uri-schemes "msteams://")

(defun org-outlook-timestamp-to-list (timestamp)
  "Parse timestamp from Outlook API.
Timestamps are returned in local timezone due to Prefer header,
so we parse them as-is without forcing UTC conversion."
  (let ((timetuple (parse-iso8601-time-string timestamp)))
    (list (string-to-number (format-time-string "%Y" timetuple))
	  (string-to-number (format-time-string "%m" timetuple))
	  (string-to-number (format-time-string "%d" timetuple))
	  (string-to-number (format-time-string "%H" timetuple))
	  (string-to-number (format-time-string "%M" timetuple)))))

(defun org-outlook-convert-html-body (html)
  (with-temp-buffer
    (insert html)
    (let ((org-inhibit-startup t)) (org-mode))
    (call-interactively 'html2org)
    (buffer-substring-no-properties (point-min) (point-max))))
(defun attendee-list (attendees &optional responsefilter)
  (let* ((attendees (append attendees nil))
	 (selected (-filter (lambda (item)(string= responsefilter (assoc-default 'response (assoc-default 'status item)))) attendees)))
    (mapconcat (lambda (item) (concat "\"" (assoc-default 'name (assoc-default 'emailAddress item)) "\"<" (assoc-default 'address (assoc-default 'emailAddress item)) ">" )) selected " ")))

(defun org-outlook-build-element (event)
  (let* ((title (assoc-default 'subject event))
	 (start (org-outlook-timestamp-to-list (assoc-default 'dateTime (assoc-default 'start event))))
	 (end (org-outlook-timestamp-to-list (assoc-default 'dateTime (assoc-default 'end event))))
	 (outlook-id (assoc-default 'id event)) ; was 'id
	 (id (secure-hash 'sha256 outlook-id))
	 (todo-state (if (not (eq :json-false (assoc-default 'isCancelled event))) "CANCELLED"
                       (if (string=  (assoc-default 'response (assoc-default 'responseStatus event)) "notResponded")
			   "REQUEST"
			 "MEETING")))
	 (locationlist (assoc-default 'displayName (assoc-default 'location event)))
	 (location (if locationlist (s-replace-regexp "\\\n" ","  locationlist) "none"))
	 (url (or (assoc-default 'webLink event) "none"))
	 (teams (or (assoc-default 'joinUrl (assoc-default 'onlineMeeting event)) "none"))
	 (accepted (attendee-list (assoc-default 'attendees event) "accepted"))
	 (declined (attendee-list (assoc-default 'attendees event) "decline"))
  	 (changekey (or (assoc-default 'changeKey event) "none"))
	 (no-response (attendee-list (assoc-default 'attendees event) "none"))
	 (categories (append (assoc-default 'categories event) nil))
	 (timestamp (->> (org-ml-build-timestamp! start :active t :end end) (org-ml-to-trimmed-string)) )
	 (html-body (or (assoc-default 'content (assoc-default 'body event)) "None"))
         (body (org-outlook-convert-html-body html-body)))
    (if title (->>
	       (org-ml-build-headline! :level 2 :title-text title :tags categories :todo-keyword todo-state :section-children
				       (list
					(org-ml-build-property-drawer!
       					 `("ID" ,id)
					 `("OUTLOOK-ID" ,outlook-id)
					 `("LOCATION" ,location)
					 `("URL" ,url)
					 `("ACCEPTED" ,accepted)
					 `("DECLINED" ,declined)
					 `("NO-RESPONSE" ,no-response)
					 `("TEAMSURL" ,teams)
		           		 `("CHANGEKEY" ,changekey)
					 ;`("MEETING-TIME" ,timestamp)
					 )
                                        (org-ml-build-paragraph! timestamp)
					(org-ml-build-clock! start :end end)
					(org-ml-build-paragraph body)))
	       (org-ml-to-trimmed-string))
      "\n")
    ))

(defun org-outlook-create-attendees-list (attendees-string)
  (->> attendees-string
       split-string
       (mapcar (lambda (x) `("emailAddress" . (("address" . ,x)))))
       (vector)))

(defun event-title (event)
  (assoc-default 'subject event))

(defun org-outlook-teams-installed-p ()
  "Check if Microsoft Teams is installed.
Returns non-nil if Teams desktop app is detected."
  (or
   ;; Check if 'teams' command exists in PATH
   (executable-find "teams")
   ;; Check common Teams installation paths on Linux
   (file-exists-p "/usr/bin/teams")
   (file-exists-p "/usr/local/bin/teams")
   (file-exists-p (expand-file-name "~/.local/share/applications/teams.desktop"))
   ;; Check Snap installation
   (file-exists-p "/snap/bin/teams")
   (file-exists-p "/var/lib/snapd/snap/bin/teams")
   ;; Check Flatpak installation
   (and (executable-find "flatpak")
        (eq 0 (call-process "flatpak" nil nil nil "info" "com.microsoft.Teams")))))

(defun org-outlook-join-teams-call ()
  "Join Teams meeting from current headline.
Uses msteams:// protocol if Teams desktop app is installed,
otherwise falls back to HTTPS URL for browser-based Teams."
  (interactive)
  (let* ((teams-url (->> (org-ml-parse-headline-at (point))
                         (org-ml-headline-get-node-property "TEAMSURL")))
         (use-teams-protocol (org-outlook-teams-installed-p))
         (final-url (if use-teams-protocol
                        (replace-regexp-in-string "https:" "msteams:" teams-url)
                      teams-url)))
    (if teams-url
        (progn
          (message (if use-teams-protocol
                       "Opening Teams meeting in desktop app..."
                     "Opening Teams meeting in browser (Teams app not detected)..."))
          (browse-url-xdg-open final-url))
      (message "No Teams URL found for this event"))))
(setq org-outlook-staging-file "~/.cache/outlook-staging.org")
(with-eval-after-load 'org-capture
  (defun org-outlook-capture-template ()
    "Returns `org-capture' template string for new outlook calendar event.
 See `org-capture-templates' for more information."
    (let* ((title (read-from-minibuffer "Event title: ")))
      (mapconcat #'identity
                 `(
                   ,(concat "** MEETING " title)
                   ":PROPERTIES:"
		   ":INVITEES: %^{Space separated invitees: }"
		   ":LOCATION: %^{Meeting location: }"
		   ":MEETING-TIME: %^{Specify meeting time: }T" ; TODO Fix
                   ":END:"
                   "%?\n")          ;Place the cursor here finally
                 "\n")))

  (add-to-list 'org-capture-templates
               '("o"                ;`org-capture' binding + o
                 "Outlook calendar event"
                 entry
                 (file org-outlook-staging-file)
                 (function org-outlook-capture-template)
		 :kill-buffer t
					;		  :prepare-finalize #'org-outlook-finalize-capture
		 )))

(defun org-outlook-finalize-capture ()
  (save-excursion
    (goto-char (point-min))
    (if (re-search-forward "MEETING-TIME" nil t) ; TODO Fix
	(progn
	  (goto-char (point-min))
	  (org-outlook-calendar-create-or-update-event)
					; (let ((attachments (org-outlook-get-appointment-property "ATTACHMENTS")))
					;      (if attachments (mapc attachments #'org-outlook-add-attachment)))
	  )
      (message "not an outlook event"))))

(defun org-outlook-get-prop-from-agenda (prop)
  (let* ((hdmarker (or (org-get-at-bol 'org-hd-marker)
		      (org-agenda-error)))
	 (buffer (marker-buffer hdmarker))
	 (pos (marker-position hdmarker)))
    (org-with-remote-undo buffer
      (with-current-buffer buffer
	(widen)
	(goto-char pos)
        (setq theprop (org-outlook-get-appointment-property prop))
	)))
  theprop)

(defun org-outlook-respond-to-event (method &optional comment sendresponse)
  (let* ((sendresponse (or sendresponse (read-string "Send response to organiser [Y/N]? ")) )
	 (comment (if (string= sendresponse "Y") (or comment (read-string "Message to organiser: ") "") ""))
	 (sendresponse (if (string= sendresponse "Y") t :json-false))
	 (recipients (if (string= method "forward")(read-string "Forward to: ") nil))
	 (id (or (org-outlook-get-appointment-property "OUTLOOK-ID")(org-outlook-get-prop-from-agenda "OUTLOOK-ID"))))

    (setq org-outlook--temp-event-id (or (org-outlook-get-appointment-property "ID")(org-outlook-get-prop-from-agenda "ID")))
    (setq org-outlook--response-method method)

    (request (concat org-outlook-events-create-url "/" id "/" method)
      :type "POST"
      :data (json-encode (remove nil `(,(if  (not (string= method "cancel")) `("sendResponse" . ,sendresponse))
				       ,(if recipients `("ToRecipients" . [(("emailAddress" . (("address") . ,recipients)))]))
				       ,(unless (string= comment "") `("comment" . ,comment))
				       )))

      :headers `(("Authorization" . ,(concat "Bearer " (org-outlook-bearer-token)))
		 ("Content-Type" . "application/json"))
      :parser 'buffer-string
      :error (cl-function
	      (lambda (&rest args &key error-thrown &key response &key data &allow-other-keys)
		(setq my-data data)
		(setq my-response response)
		(message "Error responding to event: %S" error-thrown)))
      :status-code '((202 . (lambda (&rest _)
			      (when (string= org-outlook--response-method "decline")
				(org-outlook-delete-by-id org-outlook--temp-event-id))
			     (message (concat "Responded to: " "event"))))

		     ))))

(defun org-outlook-agenda-goto-meeting () ; TODO Fix (if used)
  (interactive)
 (let* ((time (or (org-outlook-get-appointment-property "MEETING-TIME")
		 (org-outlook-get-prop-from-agenda "MEETING-TIME"))))

   (org-agenda nil "n")
   (org-agenda-goto-date time)))


(defun org-outlook-accept-event ()
  (interactive)
  (org-outlook-respond-to-event "accept"))

(defun org-outlook-decline-event ()
  (interactive)
  (let* ((data (request-response-data (org-outlook-respond-to-event "decline") ))

	 )
    (setq my-data data)))

(defun org-outlook-tentatively-accept-event ()
  (interactive)
  (org-outlook-respond-to-event "tentativelyAccept"))

(defun org-outlook-cancel-event () ;TODO Fix Json encode error
  (interactive)
  (org-outlook-respond-to-event "cancel"))

(defun org-outlook-forward-event () ;TODO test
  (interactive)
  (org-outlook-respond-to-event "forward"))


(defun org-outlook-add-attachment (myfile)
  (let* ((id (org-outlook-get-appointment-property "OUTLOOK-ID"))
	 (filename (file-name-nondirectory myfile))
	 (filebytes (with-temp-buffer
		      (let ((coding-system-for-read 'no-conversion)))
		      (insert-file-contents myfile)
		      (base64-encode-region (point-min)(point-max) t)
		      (buffer-string))))
    (request (concat org-outlook-events-create-url "/" id "/attachments")
      :type "POST"
      :data (json-encode `(("@odata.type" . "#microsoft.graph.fileAttachment")
			   ("name" . ,filename)
			   ("contentBytes" . ,filebytes)
			   ))

      :headers `(("Authorization" . ,(concat "Bearer " (org-outlook-bearer-token)))
		 ("Content-Type" . "application/json"))
      :parser 'json-read
      :error (cl-function
	      (lambda (&rest args &key error-thrown &key response &key data &allow-other-keys)
		(setq my-data data)
		(setq my-response response)
		(message "Error adding attachment: %S" error-thrown)))
      :status-code '((201 .(lambda (&rest _)
			     (message "Added attachment")))))))

(add-hook 'org-capture-prepare-finalize-hook #'org-outlook-finalize-capture)
(setq org-outlook-events-delta-url "https://graph.microsoft.com/v1.0/me/calendarview/delta")

(defun org-outlook-delete-by-id (id)
  (let* ((location (org-id-find id))
	 (file (car location))
	 (pos (cdr location))
	 )
    (with-current-buffer (or (get-file-buffer file)(find-file-noselect file))
      (unless (derived-mode-p 'org-mode)
        (let ((org-inhibit-startup t)) (org-mode)))
      (goto-char pos)
      (let*
	  ((element (org-element-at-point))
	   (begin (org-element-property :begin element))
	   (end (org-element-property :end element)))
	(delete-region begin end)
	(save-buffer)
	))))

(defun org-outlook-retrieve-events-delta (&optional nextlink)
  (interactive)
  (let ((next (or nextlink 0)))
    (message (concat "delta API call: " (or nextlink "First page")))
    (request (or nextlink org-outlook-events-delta-url)
      :params (if nextlink nil `(("startdatetime" . ,(concat (org-outlook-start) "T00:00:00.000Z" )) ("enddatetime" . ,(concat (org-outlook-end) "T23:59:59.000Z" ))) )
      :type "GET"
      :sync t
      :headers `(("Authorization" . ,(concat "Bearer " (org-outlook-bearer-token))))
      :parser 'json-read
      :error (cl-function
	      (lambda (&rest args &key error-thrown &allow-other-keys)
		(message "Error retrieving events: %S" error-thrown)))
      :success (cl-function
		(lambda (&key response &allow-other-keys)
		  (when response
		    (message "got response")
		    )
		  )
		))))

(defun org-outlook-last-delta ()
  (let*
      ((org-outlook-token-cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
       (token (plist-get (cdr (plstore-get org-outlook-token-cache "delta")) :delta)))
    (plstore-close org-outlook-token-cache)
    token))


(defun org-outlook-retrieve-delta-pages (&optional nextlink)
  (interactive)
  (message (concat "Retrieving: " (or nextlink "first page")))
  (let* ((nextlink (or nextlink (org-outlook-last-delta)))
	 (data (request-response-data (org-outlook-retrieve-events-delta nextlink)))
	 (events (append (cdr (nth 1 data)) nil))
	 (deltalink (assoc-default '@odata.deltaLink data))
	 (nextlink (assoc-default '@odata.nextLink data)))
    (if deltalink
	(progn
          (org-outlook-set-token-field "delta" `(:delta ,deltalink) '(:noclobber "noclobber"))
	  events)
      (append events (org-outlook-retrieve-delta-pages nextlink))
      )))

(defun org-outlook-retrieve-events (&optional skip)
  (let ((next (or skip 0)))
    (request org-outlook-events-url
      :params `(("startdatetime" . ,(concat (org-outlook-start) "T00:00:00.000Z" )) ("enddatetime" . ,(concat (org-outlook-end) "T23:59:59.000Z" )) ("$skip" . ,next))
      :type "GET"
      :sync t
      :headers `(("Authorization" . ,(concat "Bearer " (org-outlook-bearer-token)))
 		 ("Prefer" . ,(format "outlook.timezone=\"%s\"" org-outlook-local-timezone)))
      :parser 'json-read
      :error (cl-function
	      (lambda (&rest args &key error-thrown &allow-other-keys)
		(message "Error retrieving events: %S" error-thrown)))
      :success (cl-function
		(lambda (&key data &allow-other-keys)
		  (when data
		    (message "data received")
		    ))))))

(defun org-outlook-retrieve-pages (&optional skip)
  (message (concat "Retrieving: " (or (if skip (number-to-string skip) nil) "first page")))
  (let* ((data (request-response-data (org-outlook-retrieve-events skip)))
	 (events (append (cdr (nth 1 data)) nil))
	 (nextlink (assoc-default '@odata.nextLink data)))
    (if nextlink
	(append events (org-outlook-retrieve-pages (org-outlook--skip nextlink)))
      events
      )))


(defun org-outlook-insert-or-update (event)
  (let* ((outlook-id (assoc-default 'id event))
	 (removed (assoc-default '@removed event))
 	 (changekey (assoc-default 'changeKey event))
 	 (id (secure-hash 'sha256 outlook-id))
 	 (location (org-id-find id))
	 (file (car location))
	 (pos (cdr location)))
    (message "Processing event: outlook-id=%s, id=%s, pos=%s"
             (substring outlook-id 0 20) (substring id 0 20) pos)
    (if removed
	(org-outlook-delete-by-id id)
      (if pos ; the event exists
	  (with-current-buffer (or (get-file-buffer file)(find-file-noselect file))
	    (unless (derived-mode-p 'org-mode)
	      (let ((org-inhibit-startup t)) (org-mode)))
	    (goto-char pos)
;	    (message "At position %d in buffer %s (mode: %s)" pos (buffer-name) major-mode)
	    (let*
		((element (org-element-at-point))
;		 (_ (message (format "%s" element)))
		 (change (org-outlook-get-appointment-property "CHANGEKEY"))
		 (begin (org-element-property :begin element))
		 (end (org-element-property :end element)))
	     ; (message "Element parsed: begin=%s end=%s" begin end)
	      ;; Check if we found a valid element with begin/end positions
	      (if (and begin end (not (string= changekey change)))
		  (progn
		    (message (concat "Updating event: " outlook-id))
		    (delete-region begin end)
		    (insert (org-outlook-build-element event))
		    (insert "\n")
		    (save-buffer)
		    ;; Refresh org-id cache since file positions changed
		    (org-id-update-id-locations (list (buffer-file-name))))
		;; If element is invalid (stale org-id cache), treat as new event
		(when (not (and begin end))
		  (message (concat "Stale org-id cache for: " outlook-id " - reinserting"))
		  (with-current-buffer (or (get-file-buffer org-outlook-file)(find-file-noselect org-outlook-file))
		    (goto-char (point-max))
		    (insert (org-outlook-build-element event))
		    (insert "\n")
		    (save-buffer))))))
        (with-current-buffer (or (get-file-buffer org-outlook-file)(find-file-noselect org-outlook-file))
	  (unless (derived-mode-p 'org-mode)
	    (let ((org-inhibit-startup t)) (org-mode)))
	  (message "Inserting new event")
	  (goto-char (point-max))
	  (unless (bolp) (insert "\n"))
	  (when (org-at-heading-p) (outline-next-heading))
	  (insert (org-outlook-build-element event))
	  (insert "\n")
	  (save-buffer)
	  ;; Refresh org-id cache since file positions changed
	  (org-id-update-id-locations (list (buffer-file-name))))))))


(defvar org-outlook-full-sync-interval-days 7
  "Number of days between full syncs for validation.
Set to nil to disable automatic full syncs.")

(defun org-outlook-last-full-sync-time ()
  "Get timestamp of last full sync."
  (let* ((cache (plstore-open (expand-file-name org-outlook-token-cache-file)))
         (time-str (plist-get (cdr (plstore-get cache "last-full-sync")) :time)))
    (plstore-close cache)
    (when time-str
      (parse-iso8601-time-string time-str))))

(defun org-outlook-record-full-sync ()
  "Record timestamp of full sync completion."
  (org-outlook-set-token-field "last-full-sync"
                               `(:time ,(format-time-string "%Y-%m-%dT%H:%M:%S" (current-time)))
                               '()))

(defun org-outlook-init-delta-link ()
  "Initialize delta link after full sync for subsequent delta syncs."
  (message "Initializing delta link...")
  (let* ((response (org-outlook-retrieve-events-delta))
         (data (request-response-data response))
         (deltalink (assoc-default '@odata.deltaLink data)))
    (if deltalink
        (progn
          (org-outlook-set-token-field "delta" `(:delta ,deltalink) '())
          (message "Delta link initialized successfully"))
      (message "Warning: Could not initialize delta link"))))

(defun org-outlook-full-sync ()
  "Perform full calendar sync. Fetches all events in date range."
  (interactive)
  (message "Starting full sync...")
  (let ((newevents (org-outlook-retrieve-pages)))
    (message (concat "Full sync: processing " (number-to-string (length newevents)) " events"))
    (when newevents (mapcar 'org-outlook-insert-or-update newevents))
    (org-outlook-record-full-sync)
    (message "Full sync complete")))

(defun org-outlook-delta-sync-impl ()
  "Internal function to perform delta sync. Use `org-outlook-sync' instead."
  (let ((newevents (org-outlook-retrieve-delta-pages)))
    (message (concat "Delta sync: processing " (number-to-string (length newevents)) " changes"))
    (when newevents (mapcar 'org-outlook-insert-or-update newevents))
    (message "Delta sync complete")))

(defun org-outlook-sync (&optional force-full)
  "Sync Outlook calendar intelligently.
Uses delta sync by default for efficiency and speed.
Automatically performs full sync when:
  - No delta link exists (first run)
  - Delta link is stale (periodic validation)
  - Delta sync fails (error recovery)

With prefix argument FORCE-FULL, always performs full sync.

Delta sync is fast and handles:
  - New events
  - Updated events
  - Deleted events
  - Cancelled events

Full sync is slower but ensures no events are missed."
  (interactive "P")
  (message "Loading org-id cache...")
  (org-id-locations-load)
  (let ((delta-link (org-outlook-last-delta))
        (last-full-sync (org-outlook-last-full-sync-time)))
    (cond
     ;; User explicitly requested full sync
     (force-full
      (message "Full sync requested by user")
      (org-outlook-full-sync)
      (org-outlook-init-delta-link))

     ;; No delta link - need initial full sync
     ((not delta-link)
      (message "No delta link found - performing initial full sync")
      (org-outlook-full-sync)
      (org-outlook-init-delta-link))

     ;; Time for periodic full sync validation
     ((and org-outlook-full-sync-interval-days
           last-full-sync
           (time-less-p (time-add last-full-sync
                                  (days-to-time org-outlook-full-sync-interval-days))
                        (current-time)))
      (message "Periodic full sync (last was %d days ago)"
                       (/ (float-time (time-subtract (current-time) last-full-sync))
                          86400))
      (org-outlook-full-sync)
      (org-outlook-init-delta-link))

     ;; Normal delta sync with error recovery
     (t
      (message "Performing delta sync...")
      (condition-case err
          (org-outlook-delta-sync-impl)
        (error
         (message "Delta sync failed: %s - falling back to full sync" err)
         (org-outlook-full-sync)
         (org-outlook-init-delta-link))))))

    ;; Update org-id cache after any sync operation
    (message "Updating org-id cache...")
    (org-id-update-id-locations (list org-outlook-file))
    (org-id-locations-save))


(provide 'org-outlook)

;;; org-outlook.el ends here
