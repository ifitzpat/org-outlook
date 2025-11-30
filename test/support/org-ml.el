;;; org-ml stub for tests
(defvar org-ml-test-headline nil)

(defun org-ml-parse-headline-at (_pos)
  org-ml-test-headline)

(defun org-ml-headline-get-node-property (prop headline)
  (alist-get prop headline nil nil #'equal))

(defun org-ml-build-timestamp! (&rest _args)
  "Return placeholder timestamp string."
  "<2024-01-01 Mon>")

(defun org-ml-to-trimmed-string (value)
  (if (stringp value) value (format "%S" value)))

(defun org-ml-build-headline! (&rest args)
  args)

(defun org-ml-build-property-drawer! (&rest args)
  args)

(defun org-ml-build-paragraph! (&rest args)
  (string-join args " "))

(defun org-ml-build-clock! (&rest args)
  args)

(provide 'org-ml)
;;; org-ml.el ends here
