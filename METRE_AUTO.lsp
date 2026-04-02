;;; ============================================================
;;; METRE_AUTO.lsp  v8.0  — EXTENDED WITH QA + LOCATE + EXPORT
;;; MEXPORT uses SpreadsheetML XML — ZERO Excel COM dependency
;;; New: MWARN MLOCATE MNEXTERR MPREVERR TAGCOPY TAGSWAP
;;;      MLEVELORDER MRESUMEPIECE MEXPORTCSV MEXPORTERR
;;;      MLIST MZOOMTAG MISOROOM MISOLEVEL MMARK MLINKVIEW
;;;      MUNDO MSTAMP TAGHISTORY
;;; ============================================================
(vl-load-com)

(setq *MH*  2.75)
(setq *MU*  1.0)           ; mm=1000.0 | meters=1.0
(setq *MXL* "C:\\METRE\\METRE_EXPORT.xls")

;;; ── Global QA / State vars ────────────────────────────────
(setq *MERR_LIST*      nil)   ; list of handles with issues (MWARN)
(setq *MERR_IDX*       0)     ; current pointer for MNEXTERR/MPREVERR
(setq *MLEVEL_ORDER*   nil)   ; custom level sort order (MLEVELORDER)
(setq *MTAG_LAST_NM*   "")    ; last piece name  (TAGHISTORY)
(setq *MTAG_LAST_NV*   "")    ; last niveau       (TAGHISTORY)
(setq *MTAG_LAST_HT*   2.75)  ; last height       (TAGHISTORY)
(setq *MUNDO_SNAP*     nil)   ; undo snapshot: list of (handle xdata-str ...)

;;; =============================================================
;;; AutoCAD helpers
;;; =============================================================
(defun m:reg  (a) (if (null (tblsearch "APPID" a)) (regapp a)))
(defun m:lay  (e) (cdr (assoc 8 (entget e))))
(defun m:prop (ename / xd)
  (setq xd (assoc -3 (entget ename (list "METRE_DATA"))))
  (if xd (m:spl (cdr (assoc 1000 (cdr (cadr xd)))) "|") nil))

(defun m:prop-all (ename / xd res item)
  (setq xd (assoc -3 (entget ename '("METRE_DATA"))))
  (setq res '())
  (if xd
    (foreach item (cdadr xd)
      (if (= (car item) 1000)
        (setq res (append res (list (m:spl (cdr item) "|")))))))
  res)

(defun m:spl  (s d / p r)
  (setq r '())
  (while (setq p (vl-string-search d s))
    (setq r (append r (list (substr s 1 p))))
    (setq s (substr s (+ p 2))))
  (append r (list s)))
(defun m:len  (e / v)
  (setq v (vl-catch-all-apply 'vlax-get
             (list (vlax-ename->vla-object e) 'Length)))
  (if (vl-catch-all-error-p v) 0.0 (/ v *MU*)))
(defun m:area (e / v)
  (setq v (vl-catch-all-apply 'vlax-get
             (list (vlax-ename->vla-object e) 'Area)))
  (if (vl-catch-all-error-p v) 0.0 (/ v (* *MU* *MU*))))
(defun m:typ  (L / U)
  (setq U (strcase L))
  (cond ((wcmatch U "*MURS_INT*")   "ENDUIT_INT")
        ((wcmatch U "*MURS_EXT*")   "ENDUIT_EXT")
        ((wcmatch U "*CLOIS_INT*")  "CLOIS_INT")
        ((wcmatch U "*CLOIS_EXT*")  "CLOIS_EXT")
        ((wcmatch U "*DALLAGE*")    "DALLAGE")
        ((wcmatch U "*ENDUIT_INT*") "ENDUIT_INT")
        ((wcmatch U "*ENDUIT_EXT*") "ENDUIT_EXT")
        ((wcmatch U "*ETANCH*")     "ETANCH")
        ((wcmatch U "*PORTE*")      "PORTE")
        ((wcmatch U "*FENETRE*")    "FENETRE")
        ((wcmatch U "*FEN*")        "FENETRE")
        (T "LG")))

;;; =============================================================
;;; XML / SpreadsheetML helpers
;;; =============================================================

;;; Number to string — always uses "." (no locale issues)
(defun x:n (v)
  (vl-string-subst "." "," (rtos v 2 2)))

;;; Escape XML special characters
(defun x:e (s)
  (if (null s) (setq s ""))
  (setq s (vl-string-subst "&amp;" "&" s))
  (setq s (vl-string-subst "&lt;"  "<" s))
  (setq s (vl-string-subst "&gt;"  ">" s))
  s)

;;; Cell with String value  (st="" for no style)
(defun x:cs (v st)
  (strcat "<Cell" (if (/= st "") (strcat " ss:StyleID=\"" st "\"") "")
          "><Data ss:Type=\"String\">" (x:e v) "</Data></Cell>"))

;;; Cell with Number value
(defun x:cn (v st)
  (strcat "<Cell" (if (/= st "") (strcat " ss:StyleID=\"" st "\"") "")
          "><Data ss:Type=\"Number\">" (x:n v) "</Data></Cell>"))

;;; Empty cell
(defun x:ce () "<Cell/>")

;;; Write file header
(defun x:hdr (f)
  (write-line "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" f)
  (write-line "<?mso-application progid=\"Excel.Sheet\"?>" f)
  (write-line "<Workbook" f)
  (write-line " xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"" f)
  (write-line " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"" f)
  (write-line " xmlns:x=\"urn:schemas-microsoft-com:office:excel\">" f))

;;; Write styles
(defun x:styles (f)
  (write-line "<Styles>" f)
  ;; Default
  (write-line "<Style ss:ID=\"s0\"/>" f)
  ;; Title bar
  (write-line "<Style ss:ID=\"sT\">" f)
  (write-line " <Font ss:Bold=\"1\" ss:Color=\"#FFFFFF\" ss:Size=\"11\"/>" f)
  (write-line " <Interior ss:Color=\"#0070C0\" ss:Pattern=\"Solid\"/>" f)
  (write-line "</Style>" f)
  ;; Column header
  (write-line "<Style ss:ID=\"sH\">" f)
  (write-line " <Font ss:Bold=\"1\" ss:Color=\"#FFFFFF\" ss:Size=\"9\"/>" f)
  (write-line " <Interior ss:Color=\"#0070C0\" ss:Pattern=\"Solid\"/>" f)
  (write-line " <Alignment ss:Horizontal=\"Center\" ss:WrapText=\"1\"/>" f)
  (write-line "</Style>" f)
  ;; Niveau separator
  (write-line "<Style ss:ID=\"sN\">" f)
  (write-line " <Font ss:Bold=\"1\"/>" f)
  (write-line " <Interior ss:Color=\"#BDD7EE\" ss:Pattern=\"Solid\"/>" f)
  (write-line "</Style>" f)
  ;; Normal number
  (write-line "<Style ss:ID=\"sV\"><NumberFormat ss:Format=\"0.00\"/></Style>" f)
  ;; Alternating row
  (write-line "<Style ss:ID=\"sA\">" f)
  (write-line " <Interior ss:Color=\"#F2F2F2\" ss:Pattern=\"Solid\"/>" f)
  (write-line " <NumberFormat ss:Format=\"0.00\"/>" f)
  (write-line "</Style>" f)
  ;; Deduction row (red)
  (write-line "<Style ss:ID=\"sR\">" f)
  (write-line " <Interior ss:Color=\"#FF9999\" ss:Pattern=\"Solid\"/>" f)
  (write-line " <Font ss:Color=\"#CC0000\"/>" f)
  (write-line " <NumberFormat ss:Format=\"0.00\"/>" f)
  (write-line "</Style>" f)
  ;; Total row
  (write-line "<Style ss:ID=\"sTT\">" f)
  (write-line " <Font ss:Bold=\"1\" ss:Color=\"#FFFFFF\"/>" f)
  (write-line " <Interior ss:Color=\"#0070C0\" ss:Pattern=\"Solid\"/>" f)
  (write-line " <NumberFormat ss:Format=\"0.00\"/>" f)
  (write-line "</Style>" f)
  (write-line "</Styles>" f))

;;; Write RESUME GENERAL worksheet
(defun x:resume (f summary / gS gL gM gP rr alt st)
  (write-line "<Worksheet ss:Name=\"RESUME GENERAL\">" f)
  (write-line "<Table>" f)
  (write-line "<Column ss:Width=\"180\"/>" f)
  (write-line "<Column ss:Width=\"90\"/>" f)
  (write-line "<Column ss:Width=\"100\"/>" f)
  (write-line "<Column ss:Width=\"100\"/>" f)
  (write-line "<Column ss:Width=\"115\"/>" f)
  (write-line "<Column ss:Width=\"115\"/>" f)
  ;; Title
  (write-line "<Row ss:Height=\"20\">" f)
  (write-line "<Cell ss:MergeAcross=\"5\" ss:StyleID=\"sT\"><Data ss:Type=\"String\">RESUME GENERAL DU METRE</Data></Cell>" f)
  (write-line "</Row>" f)
  ;; Headers
  (write-line "<Row ss:Height=\"32\">" f)
  (foreach h '("LAYER / TYPE" "NB ELEMENTS" "SURF TOTAL m2"
               "LINEA TOTAL m" "ENDUIT MUR m2" "ENDUIT PLAF m2")
    (write-line (x:cs h "sH") f))
  (write-line "</Row>" f)
  ;; Data
  (setq rr 0  gS 0.0  gL 0.0  gM 0.0  gP 0.0)
  (foreach res summary
    (setq alt (= (rem rr 2) 0)  st (if alt "sA" "sV"))
    (write-line "<Row>" f)
    (write-line (x:cs (nth 0 res) (if alt "sA" "s0")) f)
    (write-line (x:cn (nth 5 res) st) f)
    (write-line (if (/= (nth 1 res) 0.0) (x:cn (nth 1 res) st) (x:ce)) f)
    (write-line (if (/= (nth 2 res) 0.0) (x:cn (nth 2 res) st) (x:ce)) f)
    (write-line (if (/= (nth 3 res) 0.0) (x:cn (nth 3 res) st) (x:ce)) f)
    (write-line (if (/= (nth 4 res) 0.0) (x:cn (nth 4 res) st) (x:ce)) f)
    (write-line "</Row>" f)
    (setq gS (+ gS (nth 1 res))  gL (+ gL (nth 2 res))
          gM (+ gM (nth 3 res))  gP (+ gP (nth 4 res)))
    (setq rr (1+ rr)))
  ;; Blank row
  (write-line "<Row/>" f)
  ;; Grand total
  (write-line "<Row ss:Height=\"18\">" f)
  (write-line (x:cs "GRAND TOTAL" "sTT") f)
  (write-line (x:ce) f)
  (write-line (if (/= gS 0.0) (x:cn gS "sTT") (x:ce)) f)
  (write-line (if (/= gL 0.0) (x:cn gL "sTT") (x:ce)) f)
  (write-line (if (/= gM 0.0) (x:cn gM "sTT") (x:ce)) f)
  (write-line (if (/= gP 0.0) (x:cn gP "sTT") (x:ce)) f)
  (write-line "</Row>" f)
  (write-line "</Table>" f)
  (write-line "</Worksheet>" f))

;;; Write ONE layer worksheet
;;; Returns (lname totS totL totM totP nbRows)
(defun x:ws (f lname rows / nom niv typ nbr sv lv hv mv pv
               totS totL totM totP prevNiv r alt st)
  (write-line (strcat "<Worksheet ss:Name=\"" (x:e (substr lname 1 31)) "\">") f)
  (write-line "<Table>" f)
  (write-line "<Column ss:Width=\"150\"/>" f)
  (write-line "<Column ss:Width=\"55\"/>" f)
  (write-line "<Column ss:Width=\"85\"/>" f)
  (write-line "<Column ss:Width=\"40\"/>" f)
  (write-line "<Column ss:Width=\"75\"/>" f)
  (write-line "<Column ss:Width=\"75\"/>" f)
  (write-line "<Column ss:Width=\"55\"/>" f)
  (write-line "<Column ss:Width=\"105\"/>" f)
  (write-line "<Column ss:Width=\"105\"/>" f)
  ;; Title
  (write-line "<Row ss:Height=\"20\">" f)
  (write-line (strcat "<Cell ss:MergeAcross=\"8\" ss:StyleID=\"sT\"><Data ss:Type=\"String\">METRE  |  "
                      (x:e lname) "</Data></Cell>") f)
  (write-line "</Row>" f)
  ;; Headers
  (write-line "<Row ss:Height=\"30\">" f)
  (foreach h '("NOM PIECE" "NIVEAU" "TYPE" "NBR"
               "SURF(m2)" "LINEA(m)" "HAUT(m)"
               "ENDUIT MUR(m2)" "ENDUIT PLAF(m2)")
    (write-line (x:cs h "sH") f))
  (write-line "</Row>" f)
  ;; Data
  (setq totS 0.0  totL 0.0  totM 0.0  totP 0.0  prevNiv ""  r 0)
  (foreach row rows
    (setq nom (nth 0 row)  niv (nth 1 row)  typ (nth 2 row)
          nbr (nth 3 row)
          sv  (atof (nth 4 row))  lv  (atof (nth 5 row))
          hv  (atof (nth 6 row))  mv  (atof (nth 7 row))
          pv  (atof (nth 8 row)))
    ;; Niveau separator
    (if (/= niv prevNiv)
      (progn
        (write-line "<Row>" f)
        (write-line (strcat "<Cell ss:MergeAcross=\"8\" ss:StyleID=\"sN\">"
                    "<Data ss:Type=\"String\">  Niveau : " (x:e niv) "</Data></Cell>") f)
        (write-line "</Row>" f)
        (setq prevNiv niv)))
    ;; Row style
    (setq alt (= (rem r 2) 0))
    (if (= (atof nbr) -1.0)
      (setq st "sR")
      (setq st (if alt "sA" "sV")))
    ;; Data row
    (write-line "<Row>" f)
    (write-line (x:cs (if (= st "sR") typ nom) (if (= st "sR") "sR" (if alt "sA" "s0"))) f)
    (write-line (x:cs niv (if (= st "sR") "sR" (if alt "sA" "s0"))) f)
    (write-line (x:cs typ (if (= st "sR") "sR" (if alt "sA" "s0"))) f)
    (write-line (x:cn (atof nbr) st) f)
    (write-line (if (/= sv 0.0) (x:cn sv st) (x:ce)) f)
    (write-line (if (/= lv 0.0) (x:cn lv st) (x:ce)) f)
    (write-line (if (/= hv 0.0) (x:cn hv st) (x:ce)) f)
    (write-line (if (/= mv 0.0) (x:cn mv st) (x:ce)) f)
    (write-line (if (/= pv 0.0) (x:cn pv st) (x:ce)) f)
    (write-line "</Row>" f)
    ;; If ETANCH → add contre receveur and deboredement sub-rows
    (if (= typ "ETANCH")
      (progn
        ;; contre receveur sub-row
        (write-line "<Row>" f)
        (write-line (x:cs "contre receveur" "sV") f)
        (write-line (x:cs niv "sV") f)
        (write-line (x:cs typ "sV") f)
        (write-line (x:cn 1.0 "sV") f)
        (write-line (x:ce) f)
        (write-line (x:cn 2.0 "sV") f)
        (write-line (x:cn 1.8 "sV") f)
        (write-line (x:ce) f)
        (write-line (x:ce) f)
        (write-line "</Row>" f)
        ;; deboredement sub-row
        (write-line "<Row>" f)
        (write-line (x:cs "deboredement" "sV") f)
        (write-line (x:cs niv "sV") f)
        (write-line (x:cs typ "sV") f)
        (write-line (x:cn 1.0 "sV") f)
        (write-line (x:ce) f)
        (write-line (x:cn 0.9 "sV") f)
        (write-line (x:cn 0.3 "sV") f)
        (write-line (x:ce) f)
        (write-line (x:ce) f)
        (write-line "</Row>" f)
        ;; Blank separator row after ETANCH group
        (write-line "<Row/>" f)))
    (setq totS (+ totS sv) totL (+ totL lv)
          totM (+ totM mv) totP (+ totP pv))
    (setq r (1+ r)))
  ;; Blank row
  (write-line "<Row/>" f)
  ;; Total row
  (write-line "<Row ss:Height=\"18\">" f)
  (write-line (x:cs "TOTAL" "sTT") f)
  (write-line (x:ce) f) (write-line (x:ce) f) (write-line (x:ce) f)
  (write-line (if (/= totS 0.0) (x:cn totS "sTT") (x:ce)) f)
  (write-line (if (/= totL 0.0) (x:cn totL "sTT") (x:ce)) f)
  (write-line (x:ce) f)
  (write-line (if (/= totM 0.0) (x:cn totM "sTT") (x:ce)) f)
  (write-line (if (/= totP 0.0) (x:cn totP "sTT") (x:ce)) f)
  (write-line "</Row>" f)
  (write-line "</Table>" f)
  (write-line "</Worksheet>" f)
  (list lname totS totL totM totP r))


;;; =============================================================
;;; HELPER: Merge rows with same NOM PIECE + NIVEAU + TYPE
;;; Sums NBR, SURF, LINEA, ENDUIT MUR, ENDUIT PLAF
;;; Keeps HEIGHT from first matching entry
;;; =============================================================
(defun m:merge-rows (rows / result row nom niv typ nbr sv lv hv mv pv key found)
  (setq result '())
  (foreach row rows
    (setq nom (nth 0 row) niv (nth 1 row) typ (nth 2 row)
          nbr (atof (nth 3 row))
          sv  (atof (nth 4 row))
          lv  (atof (nth 5 row))
          hv  (atof (nth 6 row))
          mv  (atof (nth 7 row))
          pv  (atof (nth 8 row))
          key (strcat nom "|||" niv "|||" typ)
          found nil)
    ;; PORTE and FENETRE are NEVER merged — always stay as individual deduction rows
    (if (or (wcmatch (strcase typ) "*PORTE*") (wcmatch (strcase typ) "*FENETRE*"))
      (setq result (append result
        (list (list nom niv typ
                    (rtos nbr 2 2)(rtos sv 2 2)(rtos lv 2 2)
                    (rtos hv 2 2)(rtos mv 2 2)(rtos pv 2 2)
                    key))))
      (progn
        (setq result
          (mapcar
            '(lambda (e)
               (if (= (nth 9 e) key)
                 (progn
                   (setq found T)
                   (list (nth 0 e)(nth 1 e)(nth 2 e)
                         (rtos (+ (atof (nth 3 e)) nbr) 2 2)
                         (rtos (+ (atof (nth 4 e)) sv)  2 2)
                         (rtos (+ (atof (nth 5 e)) lv)  2 2)
                         (nth 6 e)
                         (rtos (+ (atof (nth 7 e)) mv)  2 2)
                         (rtos (+ (atof (nth 8 e)) pv)  2 2)
                         (nth 9 e)))
               e))
            result))
        (if (not found)
          (setq result (append result
            (list (list nom niv typ
                        (rtos nbr 2 2)(rtos sv 2 2)(rtos lv 2 2)
                        (rtos hv 2 2)(rtos mv 2 2)(rtos pv 2 2)
                        key))))))))
  ;; Strip key (index 9) before returning
  (mapcar '(lambda (r)
             (list (nth 0 r)(nth 1 r)(nth 2 r)
                   (nth 3 r)(nth 4 r)(nth 5 r)
                   (nth 6 r)(nth 7 r)(nth 8 r)))
          result))

;;; =============================================================
;;; COMMAND: MEXPORT — Write SpreadsheetML XML file
;;; =============================================================
(defun c:MEXPORT (/ ss i e lay typ pr nm nv ht
                    nbr sv lv hv mv pv row
                    ld al res summary fout totS totL totM totP n)
  (vl-load-com)

  ;; ── Collect METRE_* objects ─────────────────────────────────
  (setq ss (ssget "X" (list (cons 0 "LWPOLYLINE,POLYLINE,LINE,ARC,CIRCLE")
                            (cons 8 "METRE_*"))))
  (if (null ss) (progn (alert "No METRE_* objects found!\nDraw on METRE_* layers first.") (exit)))

  (setq ld '() al '() i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i) lay (m:lay e) typ (m:typ lay) pr_list (m:prop-all e))
    (if (null pr_list)
      (setq pr_list (list (list "NON_TAGUE" "?" (rtos *MH* 2 2)))))

    (foreach pr pr_list
      (setq nm (nth 0 pr) nv (nth 1 pr) ht (atof (nth 2 pr)))
      (cond
        ((= typ "DALLAGE")
         (setq nbr "+1" sv (m:area e) lv 0.0 hv 0.0 mv 0.0 pv sv))
        ((= typ "ETANCH")
         (setq nbr "+1" sv (m:area e) lv (m:len e) hv 0.20 mv 0.0 pv sv))
        ((member typ '("ENDUIT_INT" "ENDUIT_EXT" "CLOIS_INT" "CLOIS_EXT"))
         (setq nbr "+1" sv (m:area e) lv (m:len e) hv ht mv (* lv hv) pv (m:area e)))
        ((member typ '("PORTE" "FENETRE"))
         (setq nbr "-1" sv 0.0 lv (m:len e) hv ht mv (* lv hv -1.0) pv 0.0))
        (T (setq nbr "+1" sv (m:area e) lv (m:len e) hv 0.0 mv 0.0 pv 0.0)))
      (setq row (list nm nv typ nbr
                      (rtos sv 2 2) (rtos lv 2 2) (rtos hv 2 2)
                      (rtos mv 2 2) (rtos pv 2 2)))
      (if (assoc lay ld)
        (setq ld (subst (cons lay (append (cdr (assoc lay ld)) (list row)))
                        (assoc lay ld) ld))
        (progn (setq ld (append ld (list (cons lay (list row)))))
               (setq al (append al (list lay))))))
    (setq i (1+ i)))

  (princ (strcat "\n  Collected: " (itoa (length al)) " layers, "
                 (itoa (sslength ss)) " objects"))

  ;; ── Distribute PORTE and FENETRE to their host layers ────────
  (setq new_ld '() doors_windows '())
  (foreach lay al
    (if (or (wcmatch lay "*PORTE*") (wcmatch lay "*FENETRE*"))
      (foreach r (cdr (assoc lay ld))
        (setq doors_windows (append doors_windows (list (cons lay r)))))
      (setq new_ld (append new_ld (list (cons lay (cdr (assoc lay ld))))))))

  (foreach dr_pair doors_windows
    (setq orig_lay (car dr_pair) dr (cdr dr_pair))
    (setq dr_nm (nth 0 dr) dr_nv (nth 1 dr) matched nil temp_ld '())
    (foreach grp new_ld
      (setq lay (car grp) rows (cdr grp) added nil)
      ;; Only add to the first matching host layer to avoid double deductions
      (if (and (not matched) (wcmatch (strcase lay) "*MURS*,*CLOIS*,*ENDUIT*"))
        (foreach r rows
          (if (and (= (nth 0 r) dr_nm) (= (nth 1 r) dr_nv))
            (setq added T))))
      (if added
        (progn
          (setq rows (append rows (list dr)) matched T)))
      (setq temp_ld (append temp_ld (list (cons lay rows)))))
    (setq new_ld temp_ld)
    (if (not matched)
      (if (assoc orig_lay new_ld)
        (setq new_ld (subst (cons orig_lay (append (cdr (assoc orig_lay new_ld)) (list dr)))
                            (assoc orig_lay new_ld) new_ld))
        (setq new_ld (append new_ld (list (cons orig_lay (list dr))))))))

  (setq ld new_ld al '())
  (foreach g ld (setq al (append al (list (car g)))))


  ;; ── Pre-compute totals for RESUME sheet ─────────────────────
  (setq summary '())
  (foreach lname al
    (setq totS 0.0  totL 0.0  totM 0.0  totP 0.0  n 0)
    (foreach row (cdr (assoc lname ld))
      (setq totS (+ totS (atof (nth 4 row)))
            totL (+ totL (atof (nth 5 row)))
            totM (+ totM (atof (nth 7 row)))
            totP (+ totP (atof (nth 8 row)))
            n (1+ n)))
    (setq summary (append summary (list (list lname totS totL totM totP n)))))

  ;; ── Create C:\METRE folder ───────────────────────────────────
  (vl-catch-all-apply 'vl-mkdir (list "C:\\METRE"))

  ;; ── Open file ───────────────────────────────────────────────
  (setq fout (open *MXL* "w"))
  (if (null fout)
    (progn (alert "ERROR: Cannot create C:\\METRE\\METRE_EXPORT.xls\nCheck that folder C:\\METRE exists.") (exit)))

  ;; ── Write XML ───────────────────────────────────────────────
  (princ "\n  Writing file...")
  (x:hdr    fout)
  (x:styles fout)

  ;; RESUME sheet first
  (x:resume fout summary)

  ;; One sheet per layer
  (foreach lname al
    (princ (strcat "\n  Sheet: " lname))
    (x:ws fout lname
      (m:merge-rows
        (vl-sort (cdr (assoc lname ld))
               '(lambda (a b)
                  (cond
                    ((not (= (cadr a) (cadr b)))
                     (< (cadr a) (cadr b)))
                    ((not (= (car a) (car b)))
                     (< (car a) (car b)))
                    (T (> (atof (cadddr a)) (atof (cadddr b))))))))))

  ;; Close XML
  (write-line "</Workbook>" fout)
  (close fout)
  (princ "\n  File closed OK.")

  ;; ── Open the file with default program (Excel) ──────────────
  (vl-catch-all-apply
    '(lambda ()
       (setq sh (vlax-create-object "WScript.Shell"))
       (vlax-invoke sh 'Run (strcat "\"" *MXL* "\"") 1 0)
       (vlax-release-object sh)))

  (alert (strcat "EXPORT COMPLETE!\n\n"
                 (itoa (length al)) " layer sheets + RESUME GENERAL\n\n"
                 "File: " *MXL* "\n\n"
                 "Open it manually if Excel didn't start:\n"
                 "C:\\METRE\\METRE_EXPORT.xls"))
  (princ))

;;; =============================================================
;;; COMMAND: MSETUP
;;; =============================================================
(defun c:MSETUP (/ doc acL ll lname lcolor lobj lf gf i fi cr up)
  (vl-load-com)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq acL (vla-get-Layers doc))
  (setq ll (list
    (list "METRE_MURS_INT"   4  "Murs int")
    (list "METRE_MURS_EXT"   5  "Murs ext")
    (list "METRE_CLOIS_INT"  3  "Cloisons int")
    (list "METRE_CLOIS_EXT"  6  "Cloisons ext")
    (list "METRE_DALLAGE"    2  "Dalle")
    (list "METRE_ENDUIT_INT" 4  "Enduit int")
    (list "METRE_ENDUIT_EXT" 5  "Enduit ext")
    (list "METRE_ETANCH"     30 "Etancheite")
    (list "METRE_LG"         7  "Lineaire")
    (list "METRE_PORTE"      1  "Portes DED")
    (list "METRE_FENETRE"    1  "Fenetres DED")))
  (setq cr 0 up 0)
  (foreach lay ll
    (setq lname (car lay) lcolor (cadr lay))
    (if (tblsearch "LAYER" lname)
      (progn (vla-put-Color (vla-Item acL lname) lcolor)
             (setq up (1+ up)) (princ (strcat "\n  [UP]  " lname)))
      (progn (setq lobj (vla-Add acL lname))
             (vla-put-Color lobj lcolor)
             (vla-put-LayerOn lobj :vlax-true)
             (vla-put-Freeze  lobj :vlax-false)
             (vla-put-Lock    lobj :vlax-false)
             (setq cr (1+ cr)) (princ (strcat "\n  [NEW] " lname)))))
  (setq lf (vla-get-LayerFilters doc))
  (setq i 0)
  (while (< i (vla-get-Count lf))
    (setq fi (vla-Item lf i))
    (if (= (strcase (vla-get-Name fi)) "METRE")
      (progn (vl-catch-all-apply 'vla-Delete (list fi)) (setq i 0))
      (setq i (1+ i))))
  (setq gf (vla-Add lf "METRE"))
  (foreach lay ll (vl-catch-all-apply '(lambda (n)(vla-Add gf n)) (list (car lay))))
  (vl-catch-all-apply 'vla-Regen (list doc acAllViewports))
  (princ (strcat "\n  DONE: " (itoa cr) " created / " (itoa up) " updated"
                 "\n  LA → filter METRE visible in left panel"))
  (princ))

;;; =============================================================
;;; COMMAND: TAG
;;; =============================================================
(defun c:TAG (/ ent e nm nv ht)
  (m:reg "METRE_DATA")
  (princ "\n[TAG] Pick element: ")
  (setq ent (entsel))
  (if ent
    (progn
      (setq e (car ent))
      (setq nm (getstring T "\n  Piece name : "))
      (setq nv (getstring T "\n  Niveau     : "))
      (setq ht (getreal (strcat "\n  Height m [" (rtos *MH* 2 2) "]: ")))
      (if (null ht) (setq ht *MH*))
      (entmod (append (entget e)
        (list (list -3 (list "METRE_DATA"
          (cons 1000 (strcat nm "|" nv "|" (rtos ht 2 3))))))))
      (progn (entupd e) (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*")) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 11)) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 256))))
      (princ (strcat "\n  OK → " nm " | " nv " | H=" (rtos ht 2 2) "m"))))
  (princ))

;;; =============================================================
;;; COMMAND: TAGM
;;; =============================================================
(defun c:TAGM (/ ss i e nm nv ht n)
  (m:reg "METRE_DATA")
  (setq nm (getstring T "\n[TAGM] Piece name : "))
  (setq nv (getstring T "       Niveau     : "))
  (setq ht (getreal (strcat "       Height m [" (rtos *MH* 2 2) "]: ")))
  (if (null ht) (setq ht *MH*))
  (princ "\n  Select objects → ")
  (setq ss (ssget) n 0 i 0)
  (if ss
    (progn
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (entmod (append (entget e)
          (list (list -3 (list "METRE_DATA"
            (cons 1000 (strcat nm "|" nv "|" (rtos ht 2 3))))))))
        (progn (entupd e) (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*")) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 11)) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 256)))) (setq i (1+ i) n (1+ n)))
      (princ (strcat "\n  OK → " (itoa n) " objects tagged → " nm " | " nv))))
  (princ))

;;; =============================================================
;;; COMMAND: MCHECK
;;; =============================================================
(defun c:MCHECK (/ ss i e cnt lay obj)
  (setq ss (ssget "X" '((8 . "METRE_*"))) cnt 0)
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i) lay (m:lay e) obj (vlax-ename->vla-object e))
        (if (null (m:prop e))
          (progn (vla-put-Color obj 1) (setq cnt (1+ cnt)))
          (if (or (wcmatch lay "*PORTE*") (wcmatch lay "*FENETRE*"))
            (vl-catch-all-apply 'vla-put-Color (list obj 11))
            (vl-catch-all-apply 'vla-put-Color (list obj 256))
          )
        )
        (setq i (1+ i)))
      (if (= cnt 0)
        (princ "\n[MCHECK] All objects tagged. Ready for MEXPORT!")
        (princ (strcat "\n[MCHECK] " (itoa cnt) " untagged (RED). They will revert to normal color when tagged.")))))
  (princ))


;;; =============================================================
;;; COMMAND: TAGLINK (Match Multiple Room Tags from Walls to Door)
;;; =============================================================
(defun c:TAGLINK (/ ss_src i_src e_src pr_list all_tags nm nv match_found t_item dht ss i n e xdata_inner)
  (m:reg "METRE_DATA")
  (princ "\n[TAGLINK] 1. Select the SOURCE object(s) (Wall/Room) to copy tag(s) FROM: ")
  (setq ss_src (ssget))
  (setq all_tags '())
  (if ss_src
    (progn
      (setq i_src 0)
      (while (< i_src (sslength ss_src))
        (setq e_src (ssname ss_src i_src))
        (setq pr_list (m:prop-all e_src))
        (if pr_list
          (foreach pr pr_list
            (setq nm (nth 0 pr) nv (nth 1 pr))
            (setq match_found nil)
            (foreach t_item all_tags
              (if (and (= (nth 0 t_item) nm) (= (nth 1 t_item) nv))
                (setq match_found T)))
            (if (not match_found)
              (setq all_tags (append all_tags (list (list nm nv)))))))
        (setq i_src (1+ i_src)))

      (if (= (length all_tags) 0)
        (princ "\nERROR: Selected source objects have no METRE_DATA tags. Use TAG first.")
        (progn
          (princ (strcat "\n Copied " (itoa (length all_tags)) " unique Room Tag(s)."))
          (setq dht (getreal "\n Enter Door/Window Height m [2.10]: "))
          (if (null dht) (setq dht 2.10))

          (princ "\n\n[TAGLINK] 2. Select DOORS or WINDOWS to apply these tags TO: ")
          (setq ss (ssget) n 0 i 0)
          (if ss
            (progn
              (while (< i (sslength ss))
                (setq e (ssname ss i))
                (setq xdata_inner (list "METRE_DATA"))
                (foreach t_item all_tags
                  (setq xdata_inner (append xdata_inner (list (cons 1000 (strcat (nth 0 t_item) "|" (nth 1 t_item) "|" (rtos dht 2 3)))))))
                (entmod (append (entget e) (list (list -3 xdata_inner))))
                (progn (entupd e) (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*")) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 11)) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 256))))
                (setq i (1+ i) n (1+ n)))
              (princ (strcat "\n SUCCESS -> " (itoa n) " doors/windows linked to " (itoa (length all_tags)) " room(s) with H=" (rtos dht 2 2) "m.")))
            (princ "\nNo objects selected.")
          )
        )
      )
    )
    (princ "\nNothing selected.")
  )
  (princ)
)


;;; =============================================================
;;; COMMAND: TAGFLOOR (Mass-update Niveau and Height only)
;;; =============================================================
(defun c:TAGFLOOR (/ nv ht ss i n e pr_list nm updated_xdata new_ht)
  (m:reg "METRE_DATA")
  (setq nv (getstring T "\n[TAGFLOOR] Enter NIVEAU for all selected objects : "))
  (setq ht (getreal (strcat "\n Enter HEIGHT m [" (rtos *MH* 2 2) "]: ")))
  (if (null ht) (setq ht *MH*))

  (princ "\n Select objects to apply this Niveau & Height -> ")
  (setq ss (ssget) n 0 i 0)
  (if ss
    (progn
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq pr_list (m:prop-all e))

        (setq updated_xdata (list "METRE_DATA"))
        (if pr_list
          ;; Update existing tags (keep existing Piece Name)
          (foreach pr pr_list
            (setq nm (nth 0 pr))
            (setq updated_xdata (append updated_xdata (list (cons 1000 (strcat nm "|" nv "|" (rtos ht 2 3)))))))
          ;; If no tag exists, create a new one with "NON_TAGUE" as piece name
          (setq updated_xdata (append updated_xdata (list (cons 1000 (strcat "NON_TAGUE|" nv "|" (rtos ht 2 3))))))
        )

        (entmod (append (entget e) (list (list -3 updated_xdata))))
        (progn (entupd e) (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*")) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 11)) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 256))))
        (setq i (1+ i) n (1+ n)))
      (princ (strcat "\n OK -> " (itoa n) " objects updated to Niveau: " nv " | H=" (rtos ht 2 2) "m.")))
    (princ "\n No objects selected.")
  )
  (princ)
)


;;; =============================================================
;;; COMMAND: TAGNAME (Mass-update Piece Name only)
;;; =============================================================
(defun c:TAGNAME (/ nm ss i n e pr_list nv ht updated_xdata)
  (m:reg "METRE_DATA")
  (setq nm (getstring T "\n[TAGNAME] Enter PIECE NAME for all selected objects : "))

  (princ "\n Select objects to apply this Piece Name -> ")
  (setq ss (ssget) n 0 i 0)
  (if ss
    (progn
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq pr_list (m:prop-all e))

        (setq updated_xdata (list "METRE_DATA"))
        (if pr_list
          ;; Update existing tags (keep existing Niveau and Height)
          (foreach pr pr_list
            (setq nv (nth 1 pr) ht (nth 2 pr))
            (setq updated_xdata (append updated_xdata (list (cons 1000 (strcat nm "|" nv "|" ht))))))
          ;; If no tag exists, create a new one with default Niveau and Height
          (setq updated_xdata (append updated_xdata (list (cons 1000 (strcat nm "|?|" (rtos *MH* 2 3))))))
        )

        (entmod (append (entget e) (list (list -3 updated_xdata))))
        (progn (entupd e) (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*")) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 11)) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 256))))
        (setq i (1+ i) n (1+ n)))
      (princ (strcat "\n OK -> " (itoa n) " objects renamed to Piece: " nm)))
    (princ "\n No objects selected.")
  )
  (princ)
)


;;; =============================================================
;;; COMMAND: TAGAUTO (Auto-tag polylines from text inside them)
;;; =============================================================
(defun c:TAGAUTO (/ ss i e el pts txt-ss txt-ent txt-str pr_list nv ht updated_xdata n)
  (vl-load-com)
  (m:reg "METRE_DATA")
  (princ "\n[TAGAUTO] Select closed Polylines (Rooms) -> ")
  (setq ss (ssget '((0 . "LWPOLYLINE"))))
  (setq n 0)
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq el (entget e))
        ;; Extract vertices from polyline
        (setq pts '())
        (foreach item el
          (if (= (car item) 10)
            (setq pts (append pts (list (cdr item))))
          )
        )
        ;; If it has at least 3 points, search for text inside/touching
        (if (> (length pts) 2)
          (progn
            ;; clear previous selection
            (setq txt-ss nil)
            (vl-catch-all-apply
              '(lambda ()
                 (setq txt-ss (ssget "CP" pts '((0 . "TEXT,MTEXT"))))
               )
            )
            (if txt-ss
              (progn
                (setq txt-ent (ssname txt-ss 0))
                ;; Get text content (handles short TEXT and simple MTEXT)
                (setq txt-str (cdr (assoc 1 (entget txt-ent))))
                ;; Clean up basic MTEXT formatting if present (e.g. \A1; \px;)
                (while (vl-string-search "\\" txt-str)
                  (setq txt-str (substr txt-str (+ 2 (vl-string-search ";" txt-str)))))
                (if (vl-string-search "{" txt-str)
                  (setq txt-str (vl-string-subst "" "{" (vl-string-subst "" "}" txt-str))))

                ;; Retain existing Niveau/Height if already tagged
                (setq pr_list (m:prop-all e))
                (setq updated_xdata (list "METRE_DATA"))
                (if pr_list
                  (foreach pr pr_list
                    (setq nv (nth 1 pr) ht (nth 2 pr))
                    (setq updated_xdata (append updated_xdata (list (cons 1000 (strcat txt-str "|" nv "|" ht))))))
                  (setq updated_xdata (append updated_xdata (list (cons 1000 (strcat txt-str "|?|" (rtos *MH* 2 3))))))
                )
                (entmod (append el (list (list -3 updated_xdata))))
                (progn (entupd e) (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*")) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 11)) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) 256))))
                (setq n (1+ n))
              )
            )
          )
        )
        (setq i (1+ i))
      )
      (princ (strcat "\n OK -> " (itoa n) " rooms automatically tagged with text found inside!"))
    )
    (princ "\n No polylines selected.")
  )
  (princ)
)


;;; =============================================================
;;; COMMAND: UNTAGLINK (Remove METRE_DATA tags from objects)
;;; =============================================================
(defun c:UNTAGLINK (/ ss i n e obj)
  (vl-load-com)
  (princ "\n[UNTAGLINK] Select objects to UNTAG (remove data) -> ")
  (setq ss (ssget) n 0 i 0)
  (if ss
    (progn
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq obj (vlax-ename->vla-object e))
        ;; Remove XData
        (entmod (list (cons -1 e) (list -3 (list "METRE_DATA"))))
        ;; Turn it RED to visually confirm it's untagged
        (vl-catch-all-apply 'vla-put-Color (list obj 1))
        (entupd e)
        (setq i (1+ i) n (1+ n)))
      (princ (strcat "\n OK -> " (itoa n) " objects UNTAGGED (Color changed to RED).")))
    (princ "\n No objects selected.")
  )
  (princ)
)



;;; =============================================================
;;; COMMAND: MEXPORTSEL — Export only selected objects
;;; =============================================================
(defun c:MEXPORTSEL (/ ss i e lay typ pr nm nv ht
                    nbr sv lv hv mv pv row
                    ld al res summary fout totS totL totM totP n dr_pair orig_lay dr dr_nm dr_nv matched temp_ld grp rows added r sh)
  (vl-load-com)

  (princ "\n[MEXPORTSEL] Select objects to export: ")
  (setq ss (ssget (list (cons 0 "LWPOLYLINE,POLYLINE,LINE,ARC,CIRCLE")
                        (cons 8 "METRE_*"))))
  (if (null ss) (progn (alert "No valid METRE_* objects selected!") (exit)))

  (setq ld '() al '() i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i) lay (m:lay e) typ (m:typ lay) pr_list (m:prop-all e))
    (if (null pr_list)
      (setq pr_list (list (list "NON_TAGUE" "?" (rtos *MH* 2 2)))))

    (foreach pr pr_list
      (setq nm (nth 0 pr) nv (nth 1 pr) ht (atof (nth 2 pr)))
      (cond
        ((= typ "DALLAGE")
         (setq nbr "+1" sv (m:area e) lv 0.0 hv 0.0 mv 0.0 pv sv))
        ((= typ "ETANCH")
         (setq nbr "+1" sv (m:area e) lv (m:len e) hv 0.20 mv 0.0 pv sv))
        ((member typ '("ENDUIT_INT" "ENDUIT_EXT" "CLOIS_INT" "CLOIS_EXT"))
         (setq nbr "+1" sv (m:area e) lv (m:len e) hv ht mv (* lv hv) pv (m:area e)))
        ((member typ '("PORTE" "FENETRE"))
         (setq nbr "-1" sv 0.0 lv (m:len e) hv ht mv (* lv hv -1.0) pv 0.0))
        (T (setq nbr "+1" sv (m:area e) lv (m:len e) hv 0.0 mv 0.0 pv 0.0)))
      (setq row (list nm nv typ nbr
                      (rtos sv 2 2) (rtos lv 2 2) (rtos hv 2 2)
                      (rtos mv 2 2) (rtos pv 2 2)))
      (if (assoc lay ld)
        (setq ld (subst (cons lay (append (cdr (assoc lay ld)) (list row)))
                        (assoc lay ld) ld))
        (progn (setq ld (append ld (list (cons lay (list row)))))
               (setq al (append al (list lay))))))
    (setq i (1+ i)))

  (princ (strcat "\n  Collected: " (itoa (length al)) " layers, "
                 (itoa (sslength ss)) " objects"))

  ;; ── Distribute PORTE and FENETRE to their host layers ────────
  (setq new_ld '() doors_windows '())
  (foreach lay al
    (if (or (wcmatch lay "*PORTE*") (wcmatch lay "*FENETRE*"))
      (foreach r (cdr (assoc lay ld))
        (setq doors_windows (append doors_windows (list (cons lay r)))))
      (setq new_ld (append new_ld (list (cons lay (cdr (assoc lay ld))))))))

  (foreach dr_pair doors_windows
    (setq orig_lay (car dr_pair) dr (cdr dr_pair))
    (setq dr_nm (nth 0 dr) dr_nv (nth 1 dr) matched nil temp_ld '())
    (foreach grp new_ld
      (setq lay (car grp) rows (cdr grp) added nil)
      (if (and (not matched) (wcmatch (strcase lay) "*MURS*,*CLOIS*,*ENDUIT*"))
        (foreach r rows
          (if (and (= (nth 0 r) dr_nm) (= (nth 1 r) dr_nv))
            (setq added T))))
      (if added
        (progn
          (setq rows (append rows (list dr)) matched T)))
      (setq temp_ld (append temp_ld (list (cons lay rows)))))
    (setq new_ld temp_ld)
    (if (not matched)
      (if (assoc orig_lay new_ld)
        (setq new_ld (subst (cons orig_lay (append (cdr (assoc orig_lay new_ld)) (list dr)))
                            (assoc orig_lay new_ld) new_ld))
        (setq new_ld (append new_ld (list (cons orig_lay (list dr))))))))

  (setq ld new_ld al '())
  (foreach g ld (setq al (append al (list (car g)))))

  ;; ── Pre-compute totals for RESUME sheet ─────────────────────
  (setq summary '())
  (foreach lname al
    (setq totS 0.0  totL 0.0  totM 0.0  totP 0.0  n 0)
    (foreach row (cdr (assoc lname ld))
      (setq totS (+ totS (atof (nth 4 row)))
            totL (+ totL (atof (nth 5 row)))
            totM (+ totM (atof (nth 7 row)))
            totP (+ totP (atof (nth 8 row)))
            n (1+ n)))
    (setq summary (append summary (list (list lname totS totL totM totP n)))))

  ;; ── Create C:\METRE folder ───────────────────────────────────
  (vl-catch-all-apply 'vl-mkdir (list "C:\\METRE"))

  ;; ── Open file ───────────────────────────────────────────────
  (setq fout (open *MXL* "w"))
  (if (null fout)
    (progn (alert "ERROR: Cannot create C:\\METRE\\METRE_EXPORT.xls\nCheck that folder C:\\METRE exists.") (exit)))

  ;; ── Write XML ───────────────────────────────────────────────
  (princ "\n  Writing file...")
  (x:hdr    fout)
  (x:styles fout)
  (x:resume fout summary)

  (foreach lname al
    (princ (strcat "\n  Sheet: " lname))
    (x:ws fout lname
      (m:merge-rows
        (vl-sort (cdr (assoc lname ld))
               '(lambda (a b)
                  (cond
                    ((not (= (cadr a) (cadr b)))
                     (< (cadr a) (cadr b)))
                    ((not (= (car a) (car b)))
                     (< (car a) (car b)))
                    (T (> (atof (cadddr a)) (atof (cadddr b))))))))))

  (write-line "</Workbook>" fout)
  (close fout)
  (princ "\n  File closed OK.")

  (vl-catch-all-apply
    '(lambda ()
       (setq sh (vlax-create-object "WScript.Shell"))
       (vlax-invoke sh 'Run (strcat "\"" *MXL* "\"") 1 0)
       (vlax-release-object sh)))

  (alert (strcat "PARTIAL EXPORT COMPLETE!\n\n"
                 (itoa (length al)) " layer sheets + RESUME GENERAL\n\n"
                 "File: " *MXL*))
  (princ)
)

;;; =============================================================
;;; HELPER: Level sort index — uses *MLEVEL_ORDER* if defined
;;; =============================================================
(defun m:level-idx (nv / i)
  (if *MLEVEL_ORDER*
    (progn
      (setq i 0)
      (while (and (< i (length *MLEVEL_ORDER*))
                  (/= (nth i *MLEVEL_ORDER*) nv))
        (setq i (1+ i)))
      i)   ; returns (length *MLEVEL_ORDER*) if not found → goes last
    0))    ; without custom order all levels sort equal (stable by string elsewhere)

;;; =============================================================
;;; HELPER: colour an entity safely
;;; =============================================================
(defun m:col (e c) (vl-catch-all-apply 'vla-put-Color (list (vlax-ename->vla-object e) c)))

;;; =============================================================
;;; HELPER: save UNDO snapshot of a selection set
;;; =============================================================
(defun m:snap-ss (ss / i e handle snap)
  (setq snap '() i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i))
    (setq handle (cdr (assoc 5 (entget e))))
    (setq pr_list (m:prop-all e))
    (setq snap (append snap (list (cons handle pr_list))))
    (setq i (1+ i)))
  snap)

;;; =============================================================
;;; COMMAND: MUNDO — Take snapshot / Restore last snapshot
;;;   Run with no argument: prompts SNAPSHOT or RESTORE
;;; =============================================================
(defun c:MUNDO (/ choice handle ename pr_list xd new_xd)
  (m:reg "METRE_DATA")
  (initget "Snap Restore")
  (setq choice (getkword "\n[MUNDO] [Snap/Restore]: "))
  (cond
    ((= choice "Snap")
     (setq tmp_ss (ssget "X" '((8 . "METRE_*"))))
     (if tmp_ss
       (progn
         (setq *MUNDO_SNAP* (m:snap-ss tmp_ss))
         (princ (strcat "\n  Snapshot saved: " (itoa (length *MUNDO_SNAP*)) " objects.")))
       (princ "\n  No METRE_* objects to snapshot.")))
    ((= choice "Restore")
     (if (null *MUNDO_SNAP*)
       (princ "\n  No snapshot — run MUNDO Snap first.")
       (progn
         (setq cnt 0)
         (foreach entry *MUNDO_SNAP*
           (setq handle (car entry))
           (setq pr_list (cdr entry))
           (setq ename (handent handle))
           (if ename
             (progn
               (setq new_xd (list "METRE_DATA"))
               (foreach pr pr_list
                 (setq new_xd (append new_xd
                   (list (cons 1000 (strcat (nth 0 pr) "|" (nth 1 pr) "|" (nth 2 pr)))))))
               (entmod (append (entget ename) (list (list -3 new_xd))))
               (entupd ename)
               (setq cnt (1+ cnt)))))
         (princ (strcat "\n  Restored " (itoa cnt) " objects from snapshot."))))))
  (princ))

;;; =============================================================
;;; COMMAND: MLEVELORDER — Set custom level sort order
;;; =============================================================
(defun c:MLEVELORDER (/ levels inp)
  (princ "\n[MLEVELORDER] Enter levels in order (blank to finish).")
  (princ "\n  Example: SS2 SS1 RDC R+1 R+2 TOIT\n")
  (setq levels '())
  (setq inp (getstring T (strcat "\n  Level " (itoa (1+ (length levels))) ": ")))
  (while (/= inp "")
    (setq levels (append levels (list inp)))
    (setq inp (getstring T (strcat "  Level " (itoa (1+ (length levels))) ": "))))
  (if levels
    (progn
      (setq *MLEVEL_ORDER* levels)
      (princ "\n  Order set: ")
      (foreach lv *MLEVEL_ORDER* (princ (strcat lv " > "))))
    (progn
      (setq *MLEVEL_ORDER* nil)
      (princ "\n  Order cleared — alphabetical will be used.")))
  (princ))

;;; =============================================================
;;; COMMAND: TAGHISTORY — Show / set last-used defaults
;;; =============================================================
(defun c:TAGHISTORY (/)
  (princ (strcat "\n[TAGHISTORY] Last values:"
                 "\n  Piece : " *MTAG_LAST_NM*
                 "\n  Niveau: " *MTAG_LAST_NV*
                 "\n  Height: " (rtos *MTAG_LAST_HT* 2 2) "m"))
  (initget "Keep Edit")
  (setq ch (getkword "\n  [Keep/Edit]: "))
  (if (= ch "Edit")
    (progn
      (setq nm (getstring T (strcat "\n  Piece [" *MTAG_LAST_NM* "]: ")))
      (if (/= nm "") (setq *MTAG_LAST_NM* nm))
      (setq nv (getstring T (strcat "  Niveau [" *MTAG_LAST_NV* "]: ")))
      (if (/= nv "") (setq *MTAG_LAST_NV* nv))
      (setq ht (getreal (strcat "  Height [" (rtos *MTAG_LAST_HT* 2 2) "]: ")))
      (if ht (setq *MTAG_LAST_HT* ht))
      (princ "\n  Defaults updated.")))
  (princ))

;;; =============================================================
;;; COMMAND: TAGCOPY — Copy XData tag from one object to others
;;; =============================================================
(defun c:TAGCOPY (/ src_ent src_e pr_list xd_inner ss i e n)
  (m:reg "METRE_DATA")
  (princ "\n[TAGCOPY] Pick SOURCE object to copy tag FROM: ")
  (setq src_ent (entsel))
  (if (null src_ent) (progn (princ "\n  Nothing selected.") (princ) (exit)))
  (setq src_e (car src_ent))
  (setq pr_list (m:prop-all src_e))
  (if (null pr_list) (progn (princ "\n  Source has no tag. Use TAG first.") (princ) (exit)))
  (princ "\n  Tag(s) to copy:")
  (foreach pr pr_list
    (princ (strcat "\n    " (nth 0 pr) " | " (nth 1 pr) " | H=" (nth 2 pr) "m")))
  (princ "\n\n  Select DESTINATION objects -> ")
  (setq ss (ssget) n 0 i 0)
  (if ss
    (progn
      (setq xd_inner (list "METRE_DATA"))
      (foreach pr pr_list
        (setq xd_inner (append xd_inner
          (list (cons 1000 (strcat (nth 0 pr) "|" (nth 1 pr) "|" (nth 2 pr)))))))
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (entmod (append (entget e) (list (list -3 xd_inner))))
        (entupd e)
        (if (or (wcmatch (m:lay e) "*PORTE*") (wcmatch (m:lay e) "*FENETRE*"))
          (m:col e 11) (m:col e 256))
        (setq i (1+ i) n (1+ n)))
      (princ (strcat "\n  OK -> " (itoa n) " objects updated."))))
  (princ))

;;; =============================================================
;;; COMMAND: TAGSWAP — Replace one room name with another
;;; =============================================================
(defun c:TAGSWAP (/ old_nm new_nm nv_filt ss i e pr_list new_xd nm nv ht n changed)
  (m:reg "METRE_DATA")
  (setq old_nm  (getstring T "\n[TAGSWAP] Old room name to replace  : "))
  (setq new_nm  (getstring T "          New room name              : "))
  (setq nv_filt (getstring T "          Level filter (* = all)     : "))
  (if (= nv_filt "") (setq nv_filt "*"))
  (setq ss (ssget "X" '((8 . "METRE_*"))) n 0)
  (if (null ss) (progn (princ "\n  No METRE_* objects.") (princ) (exit)))
  (setq i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i))
    (setq pr_list (m:prop-all e) changed nil)
    (if pr_list
      (progn
        (setq new_xd (list "METRE_DATA"))
        (foreach pr pr_list
          (setq nm (nth 0 pr) nv (nth 1 pr) ht (nth 2 pr))
          (if (and (= nm old_nm) (wcmatch (strcase nv) (strcase nv_filt)))
            (progn
              (setq new_xd (append new_xd (list (cons 1000 (strcat new_nm "|" nv "|" ht)))))
              (setq changed T))
            (setq new_xd (append new_xd (list (cons 1000 (strcat nm "|" nv "|" ht)))))))
        (if changed
          (progn
            (entmod (append (entget e) (list (list -3 new_xd))))
            (entupd e)
            (setq n (1+ n))))))
    (setq i (1+ i)))
  (princ (strcat "\n  OK -> " (itoa n) " objects swapped: '" old_nm "' → '" new_nm "'"))
  (princ))

;;; =============================================================
;;; COMMAND: MSTAMP — Attach user / date / revision to objects
;;; =============================================================
(defun c:MSTAMP (/ user note ss i e n stamp_str)
  (m:reg "METRE_STAMP")
  (setq user (getstring T "\n[MSTAMP] Your name / initials: "))
  (setq note (getstring T "         Revision note        : "))
  (setq stamp_str (strcat user "|" note "|"
                          (menucmd "m=$(edtime,$(getvar,date),DD/MM/YYYY HH:MM:SS)")))
  (princ "\n  Select objects to stamp -> ")
  (setq ss (ssget) n 0 i 0)
  (if ss
    (progn
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (entmod (append (entget e)
          (list (list -3 (list "METRE_STAMP" (cons 1000 stamp_str))))))
        (entupd e)
        (setq i (1+ i) n (1+ n)))
      (princ (strcat "\n  Stamped " (itoa n) " objects: " stamp_str))))
  (princ))

;;; =============================================================
;;; COMMAND: MLIST — Tagged inventory in command line
;;; =============================================================
(defun c:MLIST (/ ss i e handle lay typ pr_list status nb_ok nb_miss nb_multi)
  (setq ss (ssget "X" '((8 . "METRE_*"))))
  (if (null ss)
    (princ "\n[MLIST] No METRE_* objects found.")
    (progn
      (setq nb_ok 0 nb_miss 0 nb_multi 0)
      (princ (strcat "\n[MLIST] " (itoa (sslength ss)) " objects on METRE_* layers\n"))
      (princ "  Handle   | Layer                 | Room             | Level   | H(m) | GeomType    | Status")
      (princ "\n  ---------|----------------------|------------------|---------|------|-------------|--------")
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq handle (cdr (assoc 5 (entget e))))
        (setq lay (m:lay e))
        (setq typ (cdr (assoc 0 (entget e))))
        (setq pr_list (m:prop-all e))
        (cond
          ((null pr_list)
           (setq status "MISSING TAG") (setq nb_miss (1+ nb_miss))
           (princ (strcat "\n  " handle " | " lay " | ---  | ---  | ---  | " typ " | " status)))
          ((> (length pr_list) 1)
           (setq status "MULTI-LINK") (setq nb_multi (1+ nb_multi))
           (foreach pr pr_list
             (princ (strcat "\n  " handle " | " lay " | " (nth 0 pr) " | " (nth 1 pr)
                            " | " (nth 2 pr) " | " typ " | " status))))
          (T
           (setq status "OK") (setq nb_ok (1+ nb_ok))
           (foreach pr pr_list
             (princ (strcat "\n  " handle " | " lay " | " (nth 0 pr) " | " (nth 1 pr)
                            " | " (nth 2 pr) " | " typ " | " status)))))
        (setq i (1+ i)))
      (princ (strcat "\n\n  SUMMARY: " (itoa nb_ok) " OK | "
                     (itoa nb_miss) " missing tag | "
                     (itoa nb_multi) " multi-link"))))
  (princ))

;;; =============================================================
;;; COMMAND: MLOCATE — Find tagged objects by room / level
;;; =============================================================
(defun c:MLOCATE (/ nm nv ss i e pr_list found_ss handle lay)
  (setq nm (getstring T "\n[MLOCATE] Room name (* for all): "))
  (setq nv (getstring T "          Level     (* for all): "))
  (if (= nm "") (setq nm "*"))
  (if (= nv "") (setq nv "*"))
  (setq ss (ssget "X" '((8 . "METRE_*"))))
  (setq found_ss (ssadd))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (foreach pr (m:prop-all e)
          (if (and (wcmatch (strcase (nth 0 pr)) (strcase nm))
                   (wcmatch (strcase (nth 1 pr)) (strcase nv)))
            (ssadd e found_ss)))
        (setq i (1+ i)))))
  (if (= (sslength found_ss) 0)
    (princ "\n[MLOCATE] No matching objects found.")
    (progn
      (princ (strcat "\n[MLOCATE] Found " (itoa (sslength found_ss)) " object(s):"))
      (command "._ZOOM" "Object" found_ss "")
      (setq i 0)
      (while (< i (sslength found_ss))
        (setq e (ssname found_ss i))
        (setq handle (cdr (assoc 5 (entget e))))
        (setq lay (m:lay e))
        (foreach pr (m:prop-all e)
          (princ (strcat "\n  Handle=" handle " | Layer=" lay
                         " | Room=" (nth 0 pr) " | Level=" (nth 1 pr)
                         " | H=" (nth 2 pr) "m")))
        (setq i (1+ i)))))
  (princ))

;;; =============================================================
;;; COMMAND: MZOOMTAG — Zoom to object by handle or room+level
;;; =============================================================
(defun c:MZOOMTAG (/ choice handle ename ss i e found_ss nm nv)
  (initget "Handle Room")
  (setq choice (getkword "\n[MZOOMTAG] Zoom by [Handle/Room]: "))
  (cond
    ((= choice "Handle")
     (setq handle (getstring T "\n  Enter handle: "))
     (setq ename (handent handle))
     (if ename
       (progn
         (command "._ZOOM" "Object" (ssadd ename) "")
         (princ (strcat "\n  Zoomed to handle " handle
                        " | Layer=" (m:lay ename))))
       (princ "\n  Handle not found in drawing.")))
    ((= choice "Room")
     (setq nm (getstring T "\n  Room name  : "))
     (setq nv (getstring T "  Level      : "))
     (setq ss (ssget "X" '((8 . "METRE_*"))))
     (setq found_ss (ssadd))
     (if ss
       (progn
         (setq i 0)
         (while (< i (sslength ss))
           (setq e (ssname ss i))
           (foreach pr (m:prop-all e)
             (if (and (wcmatch (strcase (nth 0 pr)) (strcase nm))
                      (wcmatch (strcase (nth 1 pr)) (strcase nv)))
               (ssadd e found_ss)))
           (setq i (1+ i)))))
     (if (> (sslength found_ss) 0)
       (progn
         (command "._ZOOM" "Object" found_ss "")
         (princ (strcat "\n  Zoomed to " (itoa (sslength found_ss)) " object(s).")))
       (princ "\n  No matching objects found."))))
  (princ))

;;; =============================================================
;;; COMMAND: MWARN — Intelligent issue detector
;;; Populates *MERR_LIST* for MNEXTERR / MPREVERR
;;; =============================================================
(defun c:MWARN (/ ss i e pr_list lay typ handle nm nv ht
                   cnt_notag cnt_noroom cnt_noniv cnt_badht
                   cnt_dup cnt_zero sv lv dk seen_keys issues)
  (setq ss (ssget "X" '((8 . "METRE_*"))))
  (if (null ss) (progn (princ "\n[MWARN] No METRE_* objects.") (princ) (exit)))
  (setq cnt_notag 0 cnt_noroom 0 cnt_noniv 0 cnt_badht 0
        cnt_dup 0 cnt_zero 0 issues '() *MERR_LIST* nil *MERR_IDX* 0)
  (setq i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i))
    (setq handle (cdr (assoc 5 (entget e))))
    (setq lay (m:lay e))
    (setq typ (m:typ lay))
    (setq pr_list (m:prop-all e))
    ;; — No XData
    (if (null pr_list)
      (progn
        (setq cnt_notag (1+ cnt_notag))
        (setq issues (append issues (list (list handle lay "MISSING TAG" "" ""))))
        (m:col e 1))
      (progn
        ;; — Per-tag checks
        (foreach pr pr_list
          (setq nm (nth 0 pr) nv (nth 1 pr) ht (atof (nth 2 pr)))
          (if (or (= nm "") (= nm "?") (= nm "NON_TAGUE"))
            (progn (setq cnt_noroom (1+ cnt_noroom))
                   (setq issues (append issues (list (list handle lay "EMPTY ROOM NAME" nm nv))))))
          (if (or (= nv "") (= nv "?"))
            (progn (setq cnt_noniv (1+ cnt_noniv))
                   (setq issues (append issues (list (list handle lay "EMPTY LEVEL" nm nv))))))
          (if (or (<= ht 0.0) (< ht 0.1) (> ht 15.0))
            (progn (setq cnt_badht (1+ cnt_badht))
                   (setq issues (append issues (list (list handle lay "BAD HEIGHT" nm nv)))))))
        ;; — Duplicate room+level on same object
        (setq seen_keys '())
        (foreach pr pr_list
          (setq dk (strcat (nth 0 pr) "|||" (nth 1 pr)))
          (if (member dk seen_keys)
            (setq cnt_dup (1+ cnt_dup))
            (setq seen_keys (append seen_keys (list dk)))))
        ;; — Near-zero geometry (non-deductions only)
        (if (not (or (wcmatch lay "*PORTE*") (wcmatch lay "*FENETRE*")))
          (progn
            (setq sv (m:area e) lv (m:len e))
            (if (and (< sv 0.001) (< lv 0.001))
              (progn (setq cnt_zero (1+ cnt_zero))
                     (setq issues (append issues (list (list handle lay "ZERO GEOMETRY" "" ""))))))))))
    (setq i (1+ i)))
  ;; Build unique *MERR_LIST*
  (setq *MERR_LIST* nil)
  (foreach iss issues
    (if (not (member (car iss) *MERR_LIST*))
      (setq *MERR_LIST* (append *MERR_LIST* (list (car iss))))))
  ;; Summary
  (princ "\n\n[MWARN] ═══ AUDIT REPORT ═══")
  (princ (strcat "\n  " (itoa cnt_notag)  " missing tags"))
  (princ (strcat "\n  " (itoa cnt_noroom) " empty room names"))
  (princ (strcat "\n  " (itoa cnt_noniv)  " empty levels"))
  (princ (strcat "\n  " (itoa cnt_badht)  " invalid heights"))
  (princ (strcat "\n  " (itoa cnt_dup)    " duplicate room+level pairs"))
  (princ (strcat "\n  " (itoa cnt_zero)   " near-zero geometry"))
  (princ (strcat "\n  ─────────────────────────────"))
  (princ (strcat "\n  " (itoa (length *MERR_LIST*)) " objects with issues total"))
  (if *MERR_LIST*
    (progn
      (princ "\n\n  Use MNEXTERR to step through each issue.")
      (princ "\n  Use MEXPORTERR to export error list to CSV.")))
  (princ))

;;; =============================================================
;;; HELPER: show one error entry and zoom to it
;;; =============================================================
(defun m:show-err (idx / handle ename pr_list)
  (if (or (null *MERR_LIST*) (= (length *MERR_LIST*) 0))
    (princ "\n  No errors — run MWARN first.")
    (progn
      (setq handle (nth idx *MERR_LIST*))
      (setq ename (handent handle))
      (if ename
        (progn
          (command "._ZOOM" "Object" (ssadd ename) "")
          (princ (strcat "\n  Error " (itoa (1+ idx)) " / " (itoa (length *MERR_LIST*))
                         "  |  Handle=" handle "  |  Layer=" (m:lay ename)))
          (setq pr_list (m:prop-all ename))
          (if pr_list
            (foreach pr pr_list
              (princ (strcat "\n    Room=" (nth 0 pr)
                             "  Level=" (nth 1 pr)
                             "  H=" (nth 2 pr) "m")))
            (princ "\n    *** NO TAG DATA ***")))
        (princ "\n  Object no longer exists (deleted?).")))))

;;; =============================================================
;;; COMMAND: MNEXTERR — Step forward through error list
;;; =============================================================
(defun c:MNEXTERR (/)
  (if (null *MERR_LIST*)
    (princ "\n[MNEXTERR] No error list. Run MWARN first.")
    (progn
      (if (>= *MERR_IDX* (length *MERR_LIST*)) (setq *MERR_IDX* 0))
      (m:show-err *MERR_IDX*)
      (setq *MERR_IDX* (1+ *MERR_IDX*))))
  (princ))

;;; =============================================================
;;; COMMAND: MPREVERR — Step backward through error list
;;; =============================================================
(defun c:MPREVERR (/)
  (if (null *MERR_LIST*)
    (princ "\n[MPREVERR] No error list. Run MWARN first.")
    (progn
      (setq *MERR_IDX* (1- *MERR_IDX*))
      (if (< *MERR_IDX* 0) (setq *MERR_IDX* (1- (length *MERR_LIST*))))
      (m:show-err *MERR_IDX*)))
  (princ))

;;; =============================================================
;;; COMMAND: MISOROOM — Zoom-isolate all objects matching a room
;;; =============================================================
(defun c:MISOROOM (/ nm ss i e found_ss)
  (setq nm (getstring T "\n[MISOROOM] Room name to isolate (* wildcard OK): "))
  (if (= nm "") (setq nm "*"))
  (setq ss (ssget "X" '((8 . "METRE_*"))))
  (setq found_ss (ssadd))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (foreach pr (m:prop-all e)
          (if (wcmatch (strcase (nth 0 pr)) (strcase nm))
            (ssadd e found_ss)))
        (setq i (1+ i)))))
  (if (> (sslength found_ss) 0)
    (progn
      (command "._ZOOM" "Object" found_ss "")
      (princ (strcat "\n  " (itoa (sslength found_ss)) " object(s) for room: " nm)))
    (princ "\n  No matching objects found."))
  (princ))

;;; =============================================================
;;; COMMAND: MISOLEVEL — Zoom-isolate all objects on one level
;;; =============================================================
(defun c:MISOLEVEL (/ nv ss i e found_ss)
  (setq nv (getstring T "\n[MISOLEVEL] Level to isolate (* wildcard OK): "))
  (if (= nv "") (setq nv "*"))
  (setq ss (ssget "X" '((8 . "METRE_*"))))
  (setq found_ss (ssadd))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (foreach pr (m:prop-all e)
          (if (wcmatch (strcase (nth 1 pr)) (strcase nv))
            (ssadd e found_ss)))
        (setq i (1+ i)))))
  (if (> (sslength found_ss) 0)
    (progn
      (command "._ZOOM" "Object" found_ss "")
      (princ (strcat "\n  " (itoa (sslength found_ss)) " object(s) for level: " nv)))
    (princ "\n  No matching objects found."))
  (princ))

;;; =============================================================
;;; COMMAND: MMARK — Create temporary MTEXT labels near objects
;;; =============================================================
(defun c:MMARK (/ ss i e pr_list handle lay obj cpt ins txt)
  (m:reg "METRE_DATA")
  (princ "\n[MMARK] Select objects to label (labels on layer METRE_MARKS) -> ")
  (setq ss (ssget) i 0)
  (if (null ss) (progn (princ "\n  Nothing selected.") (princ) (exit)))
  (while (< i (sslength ss))
    (setq e (ssname ss i))
    (setq pr_list (m:prop-all e))
    (setq handle (cdr (assoc 5 (entget e))))
    (setq lay (m:lay e))
    (if pr_list
      (progn
        ;; Try to get a usable insertion point
        (setq obj (vlax-ename->vla-object e))
        (setq cpt (vl-catch-all-apply 'vlax-get (list obj 'Centroid)))
        (if (vl-catch-all-error-p cpt)
          (setq cpt (vl-catch-all-apply 'vlax-get (list obj 'StartPoint))))
        (if (vl-catch-all-error-p cpt)
          (setq cpt (cdr (assoc 10 (entget e)))))
        (if cpt
          (progn
            (if (= (type cpt) 'VARIANT)
              (setq cpt (vlax-safearray->list (vlax-variant-value cpt))))
            ;; Build label text
            (setq txt "")
            (foreach pr pr_list
              (setq txt (strcat txt (nth 0 pr) "/" (nth 1 pr) "/H=" (nth 2 pr) "\P")))
            (setq txt (strcat txt "[" handle "]"))
            ;; Create MTEXT
            (entmake
              (list
                '(0 . "MTEXT")
                '(100 . "AcDbEntity")
                (cons 8 "METRE_MARKS")
                (cons 62 3)
                '(100 . "AcDbMText")
                (cons 10 cpt)
                (cons 40 0.25)
                (cons 41 3.0)
                (cons 1 txt)))))))
    (setq i (1+ i)))
  (princ "\n  Labels created on layer METRE_MARKS.")
  (princ "\n  Use ERASE and pick layer METRE_MARKS to remove them later.")
  (princ))

;;; =============================================================
;;; COMMAND: MLINKVIEW — Visualise which rooms a door/window links
;;; =============================================================
(defun c:MLINKVIEW (/ ent e lay pr_list handle wall_ss j we wpr_list found_ss)
  (princ "\n[MLINKVIEW] Pick a DOOR or WINDOW element: ")
  (setq ent (entsel))
  (if (null ent) (progn (princ "\n  Nothing.") (princ) (exit)))
  (setq e (car ent) lay (m:lay e))
  (if (not (or (wcmatch lay "*PORTE*") (wcmatch lay "*FENETRE*")))
    (princ "\n  Selected object is not on a PORTE or FENETRE layer.")
    (progn
      (setq pr_list (m:prop-all e))
      (setq handle (cdr (assoc 5 (entget e))))
      (princ (strcat "\n[MLINKVIEW] Handle=" handle " | Layer=" lay))
      (if (null pr_list)
        (princ "\n  No tags — not linked to any room. Use TAGLINK first.")
        (progn
          (princ (strcat "\n  Linked to " (itoa (length pr_list)) " room(s):"))
          (foreach pr pr_list
            (princ (strcat "\n    Room=" (nth 0 pr)
                           "  Level=" (nth 1 pr)
                           "  H=" (nth 2 pr) "m")))
          ;; Find matching wall segments
          (setq wall_ss (ssget "X" '((8 . "METRE_MURS_INT,METRE_MURS_EXT,METRE_CLOIS_INT,METRE_CLOIS_EXT,METRE_ENDUIT_INT,METRE_ENDUIT_EXT"))))
          (setq found_ss (ssadd))
          (if wall_ss
            (progn
              (setq j 0)
              (while (< j (sslength wall_ss))
                (setq we (ssname wall_ss j))
                (setq wpr_list (m:prop-all we))
                (foreach wpr wpr_list
                  (foreach pr pr_list
                    (if (and (= (nth 0 wpr) (nth 0 pr))
                             (= (nth 1 wpr) (nth 1 pr)))
                      (ssadd we found_ss))))
                (setq j (1+ j)))))
          (if (> (sslength found_ss) 0)
            (progn
              (princ (strcat "\n  " (itoa (sslength found_ss)) " matching wall segment(s) found."))
              (command "._ZOOM" "Object" found_ss ""))
            (princ "\n  No matching wall segments in drawing."))))))
  (princ))

;;; =============================================================
;;; COMMAND: MEXPORTCSV — Export all METRE_* data to CSV
;;; =============================================================
(defun c:MEXPORTCSV (/ ss i e lay typ pr_list pr nm nv ht
                       nbr sv lv hv mv pv fout fname fout2 lname)
  (vl-load-com)
  (vl-catch-all-apply 'vl-mkdir (list "C:\\METRE"))
  (setq fname "C:\\METRE\\METRE_EXPORT.csv")
  (setq fout (open fname "w"))
  (if (null fout) (progn (alert "Cannot create C:\\METRE\\METRE_EXPORT.csv") (exit)))

  ;; CSV header
  (write-line "Handle,Layer,Type,Room,Level,NBR,SURF_m2,LINEA_m,HEIGHT_m,ENDUIT_MUR_m2,ENDUIT_PLAF_m2" fout)

  (setq ss (ssget "X" (list (cons 0 "LWPOLYLINE,POLYLINE,LINE,ARC,CIRCLE")
                            (cons 8 "METRE_*"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq lay (m:lay e))
        (setq typ (m:typ lay))
        (setq pr_list (m:prop-all e))
        (if (null pr_list) (setq pr_list (list (list "NON_TAGUE" "?" (rtos *MH* 2 2)))))
        (foreach pr pr_list
          (setq nm (nth 0 pr) nv (nth 1 pr) ht (atof (nth 2 pr)))
          (cond
            ((= typ "DALLAGE")
             (setq nbr "+1" sv (m:area e) lv 0.0 hv 0.0 mv 0.0 pv sv))
            ((= typ "ETANCH")
             (setq nbr "+1" sv (m:area e) lv (m:len e) hv 0.20 mv 0.0 pv sv))
            ((member typ '("ENDUIT_INT" "ENDUIT_EXT" "CLOIS_INT" "CLOIS_EXT"))
             (setq nbr "+1" sv (m:area e) lv (m:len e) hv ht mv (* lv hv) pv (m:area e)))
            ((member typ '("PORTE" "FENETRE"))
             (setq nbr "-1" sv 0.0 lv (m:len e) hv ht mv (* lv hv -1.0) pv 0.0))
            (T (setq nbr "+1" sv (m:area e) lv (m:len e) hv 0.0 mv 0.0 pv 0.0)))
          (write-line
            (strcat (cdr (assoc 5 (entget e))) ","
                    lay "," typ ","
                    nm "," nv "," nbr ","
                    (x:n sv) "," (x:n lv) "," (x:n hv) ","
                    (x:n mv) "," (x:n pv))
            fout))
        (setq i (1+ i)))))
  (close fout)
  (princ (strcat "\n[MEXPORTCSV] Written: " fname))
  (alert (strcat "CSV EXPORT COMPLETE!\n\nFile: " fname))
  (princ))

;;; =============================================================
;;; COMMAND: MEXPORTERR — Export problem objects to CSV
;;; =============================================================
(defun c:MEXPORTERR (/ fout fname handle ename lay pr_list nm nv ht sv lv)
  (if (null *MERR_LIST*)
    (progn
      (princ "\n[MEXPORTERR] No error list. Running MWARN first...")
      (c:MWARN)))
  (if (null *MERR_LIST*)
    (progn (princ "\n  No issues found.") (princ) (exit)))
  (vl-catch-all-apply 'vl-mkdir (list "C:\\METRE"))
  (setq fname "C:\\METRE\\METRE_ERREURS.csv")
  (setq fout (open fname "w"))
  (if (null fout) (progn (alert "Cannot create C:\\METRE\\METRE_ERREURS.csv") (exit)))
  (write-line "Handle,Layer,Room,Level,Height,SURF_m2,LINEA_m,Problem,Suggested_Fix" fout)
  (foreach handle *MERR_LIST*
    (setq ename (handent handle))
    (if ename
      (progn
        (setq lay (m:lay ename))
        (setq pr_list (m:prop-all ename))
        (setq sv (x:n (m:area ename)) lv (x:n (m:len ename)))
        (if (null pr_list)
          (write-line (strcat handle "," lay ",,,,\"" sv "\",\"" lv "\",MISSING TAG,Run TAG or TAGM command") fout)
          (foreach pr pr_list
            (setq nm (nth 0 pr) nv (nth 1 pr) ht (atof (nth 2 pr)))
            (setq prob ""  fix "")
            (if (or (= nm "") (= nm "?") (= nm "NON_TAGUE"))
              (progn (setq prob "EMPTY ROOM NAME") (setq fix "Use TAGNAME to set room name")))
            (if (or (= nv "") (= nv "?"))
              (progn (setq prob (strcat prob " EMPTY LEVEL")) (setq fix (strcat fix " / Use TAGFLOOR"))))
            (if (or (<= ht 0.0) (< ht 0.1) (> ht 15.0))
              (progn (setq prob (strcat prob " BAD HEIGHT")) (setq fix (strcat fix " / Correct height"))))
            (if (= prob "")
              (progn (setq prob "GEOMETRY ISSUE") (setq fix "Check object geometry")))
            (write-line (strcat handle "," lay "," nm "," nv ","
                                (x:n ht) "," sv "," lv ","
                                "\"" prob "\",\"" fix "\"") fout))))))
  (close fout)
  (princ (strcat "\n[MEXPORTERR] Written: " fname))
  (alert (strcat "ERROR EXPORT COMPLETE!\n" (itoa (length *MERR_LIST*)) " problem objects.\n\nFile: " fname))
  (princ))

;;; =============================================================
;;; HELPER: Write RESUME PAR PIECE worksheet (room-based summary)
;;; =============================================================
(defun x:resume-piece (f rows / piece_map key nm nv typ nbr sv lv hv mv pv
                          k_list k entry rr alt st gM gP gL gS)
  (write-line "<Worksheet ss:Name=\"RESUME PAR PIECE\">" f)
  (write-line "<Table>" f)
  (write-line "<Column ss:Width=\"150\"/>" f)
  (write-line "<Column ss:Width=\"80\"/>" f)
  (write-line "<Column ss:Width=\"90\"/>" f)
  (write-line "<Column ss:Width=\"90\"/>" f)
  (write-line "<Column ss:Width=\"90\"/>" f)
  (write-line "<Column ss:Width=\"60\"/>" f)
  ;; Title
  (write-line "<Row ss:Height=\"20\">" f)
  (write-line "<Cell ss:MergeAcross=\"5\" ss:StyleID=\"sT\"><Data ss:Type=\"String\">RESUME PAR PIECE / NIVEAU</Data></Cell>" f)
  (write-line "</Row>" f)
  ;; Headers
  (write-line "<Row ss:Height=\"30\">" f)
  (foreach h '("PIECE" "NIVEAU" "LG MURS (m)" "ENDUIT MUR (m2)" "ENDUIT PLAF (m2)" "NB DED")
    (write-line (x:cs h "sH") f))
  (write-line "</Row>" f)
  ;; Aggregate by PIECE + NIVEAU
  (setq piece_map '() k_list '())
  (foreach row rows
    (setq nm (nth 0 row) nv (nth 1 row) typ (nth 2 row)
          nbr (atof (nth 3 row))
          lv  (atof (nth 5 row))
          mv  (atof (nth 7 row))
          pv  (atof (nth 8 row))
          key (strcat nm "|||" nv))
    (if (not (member key k_list))
      (progn
        (setq k_list (append k_list (list key)))
        (setq piece_map (append piece_map (list (list key nm nv 0.0 0.0 0.0 0))))))
    (setq piece_map
      (mapcar '(lambda (e)
                 (if (= (car e) key)
                   (list (nth 0 e)(nth 1 e)(nth 2 e)
                         (+ (nth 3 e) lv)
                         (+ (nth 4 e) mv)
                         (+ (nth 5 e) pv)
                         (if (= nbr -1.0) (1+ (nth 6 e)) (nth 6 e)))
                 e))
              piece_map)))
  ;; Sort by *MLEVEL_ORDER* if set, then by piece name
  (setq piece_map
    (vl-sort piece_map
      '(lambda (a b)
         (if *MLEVEL_ORDER*
           (< (m:level-idx (nth 2 a)) (m:level-idx (nth 2 b)))
           (< (strcase (nth 2 a)) (strcase (nth 2 b)))))))
  ;; Write rows
  (setq rr 0 gL 0.0 gM 0.0 gP 0.0)
  (foreach entry piece_map
    (setq alt (= (rem rr 2) 0) st (if alt "sA" "sV"))
    (write-line "<Row>" f)
    (write-line (x:cs (nth 1 entry) (if alt "sA" "s0")) f)
    (write-line (x:cs (nth 2 entry) (if alt "sA" "s0")) f)
    (write-line (x:cn (nth 3 entry) st) f)
    (write-line (x:cn (nth 4 entry) st) f)
    (write-line (x:cn (nth 5 entry) st) f)
    (write-line (x:cn (float (nth 6 entry)) st) f)
    (write-line "</Row>" f)
    (setq gL (+ gL (nth 3 entry)) gM (+ gM (nth 4 entry)) gP (+ gP (nth 5 entry)))
    (setq rr (1+ rr)))
  ;; Total
  (write-line "<Row/>" f)
  (write-line "<Row ss:Height=\"18\">" f)
  (write-line (x:cs "TOTAL" "sTT") f)
  (write-line (x:ce) f)
  (write-line (x:cn gL "sTT") f)
  (write-line (x:cn gM "sTT") f)
  (write-line (x:cn gP "sTT") f)
  (write-line (x:ce) f)
  (write-line "</Row>" f)
  (write-line "</Table>" f)
  (write-line "</Worksheet>" f))

;;; =============================================================
;;; COMMAND: MRESUMEPIECE — Export room-based summary to Excel
;;; =============================================================
(defun c:MRESUMEPIECE (/ ss i e lay typ pr_list pr nm nv ht
                          nbr sv lv hv mv pv row all_rows fout fname)
  (vl-load-com)
  (setq ss (ssget "X" (list (cons 0 "LWPOLYLINE,POLYLINE,LINE,ARC,CIRCLE")
                            (cons 8 "METRE_*"))))
  (if (null ss) (progn (alert "No METRE_* objects found!") (exit)))
  (setq all_rows '() i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i) lay (m:lay e) typ (m:typ lay))
    (setq pr_list (m:prop-all e))
    (if (null pr_list) (setq pr_list (list (list "NON_TAGUE" "?" (rtos *MH* 2 2)))))
    (foreach pr pr_list
      (setq nm (nth 0 pr) nv (nth 1 pr) ht (atof (nth 2 pr)))
      (cond
        ((member typ '("ENDUIT_INT" "ENDUIT_EXT" "CLOIS_INT" "CLOIS_EXT"))
         (setq nbr "+1" sv (m:area e) lv (m:len e) hv ht mv (* lv hv) pv (m:area e)))
        ((member typ '("PORTE" "FENETRE"))
         (setq nbr "-1" sv 0.0 lv (m:len e) hv ht mv (* lv hv -1.0) pv 0.0))
        ((= typ "DALLAGE")
         (setq nbr "+1" sv (m:area e) lv 0.0 hv 0.0 mv 0.0 pv sv))
        (T (setq nbr "+1" sv 0.0 lv (m:len e) hv 0.0 mv 0.0 pv 0.0)))
      (setq row (list nm nv typ nbr (rtos sv 2 2)(rtos lv 2 2)(rtos hv 2 2)(rtos mv 2 2)(rtos pv 2 2)))
      (setq all_rows (append all_rows (list row))))
    (setq i (1+ i)))
  (vl-catch-all-apply 'vl-mkdir (list "C:\\METRE"))
  (setq fname "C:\\METRE\\METRE_RESUME_PIECE.xls")
  (setq fout (open fname "w"))
  (if (null fout) (progn (alert "Cannot create file.") (exit)))
  (x:hdr fout)
  (x:styles fout)
  (x:resume-piece fout all_rows)
  (write-line "</Workbook>" fout)
  (close fout)
  (alert (strcat "RESUME PAR PIECE COMPLETE!\n\nFile: " fname))
  (princ))

;;; =============================================================
;;; Updated banner
;;; =============================================================
(princ "\n╔══════════════════════════════════════════════════╗")
(princ "\n║   METRE_AUTO  v8.0  —  LOADED                   ║")
(princ "\n╠══════════════════════════════════╦═══════════════╣")
(princ "\n║  SETUP / LAYERS                 ║  QA / AUDIT   ║")
(princ "\n║  MSETUP    Create layers        ║  MCHECK       ║")
(princ "\n║  MLEVELORDER  Custom sort order ║  MWARN        ║")
(princ "\n╠══════════════════════════════════╣  MNEXTERR     ║")
(princ "\n║  TAGGING                        ║  MPREVERR     ║")
(princ "\n║  TAG       Tag one element      ╠═══════════════╣")
(princ "\n║  TAGM      Mass-tag selection   ║  LOCATE       ║")
(princ "\n║  TAGFLOOR  Update niveau/height ║  MLOCATE      ║")
(princ "\n║  TAGNAME   Update piece name    ║  MZOOMTAG     ║")
(princ "\n║  TAGAUTO   Auto from text       ║  MLIST        ║")
(princ "\n║  TAGLINK   Link doors/windows   ║  MISOROOM     ║")
(princ "\n║  TAGCOPY   Copy tag to objects  ║  MISOLEVEL    ║")
(princ "\n║  TAGSWAP   Rename room globally ╠═══════════════╣")
(princ "\n║  TAGHISTORY  Show/set defaults  ║  EXPORT       ║")
(princ "\n║  UNTAGLINK Remove tags          ║  MEXPORT      ║")
(princ "\n╠══════════════════════════════════╣  MEXPORTSEL   ║")
(princ "\n║  VISUAL / UTILS                 ║  MEXPORTCSV   ║")
(princ "\n║  MMARK     Label objects        ║  MEXPORTERR   ║")
(princ "\n║  MLINKVIEW View door links      ║  MRESUMEPIECE ║")
(princ "\n║  MUNDO     Snapshot / Restore   ║               ║")
(princ "\n║  MSTAMP    User/date stamp      ║               ║")
(princ "\n╚══════════════════════════════════╩═══════════════╝")
(princ)
