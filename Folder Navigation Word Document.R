# ============================================================
# MODIOLUS MODIOLUS PROJECT
# AUTOMATIC WORD FOLDER NAVIGATION
#
# Creates a Microsoft Word document showing the complete
# folder/file hierarchy of the project.
#
# Uses:
#   R
#     -> PowerShell
#         -> Microsoft Word COM automation
#
# Microsoft Word must be installed on this computer.
#
# R version: 4.5.1
# Windows only
# ============================================================


# ============================================================
# 1. PROJECT SETTINGS
# ============================================================

project_folder <- "C:\\Users\\Paul-RubenJeandidier\\OneDrive - Atlantic TU\\Modiolus_modiolus_PaulRuben"

output_file <- file.path(
  project_folder,
  "Modiolus_Project_Folder_Navigation.docx"
)


# ============================================================
# 2. CHECK THAT THE PROJECT FOLDER EXISTS
# ============================================================

if (!dir.exists(project_folder)) {
  stop(
    paste0(
      "\nERROR: The project folder does not exist:\n",
      project_folder,
      "\n\nPlease check that the path is correct."
    )
  )
}


# ============================================================
# 3. CHECK THAT POWERSHELL IS AVAILABLE
# ============================================================

powershell <- Sys.which("powershell.exe")

if (powershell == "") {
  stop(
    "\nERROR: Windows PowerShell could not be found.\n",
    "This script requires PowerShell and Microsoft Word."
  )
}


# ============================================================
# 4. CHECK THAT MICROSOFT WORD IS INSTALLED
#
# We let PowerShell test Word's COM object.
# ============================================================

cat("\nChecking Microsoft Word installation...\n")

word_check_script <- '
try {
    $word = New-Object -ComObject Word.Application
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Output "WORD_OK"
}
catch {
    Write-Output "WORD_ERROR"
}
'

check_file <- tempfile(fileext = ".ps1")
writeLines(word_check_script, check_file, useBytes = TRUE)

word_check <- system2(
  powershell,
  args = c(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", shQuote(check_file)
  ),
  stdout = TRUE,
  stderr = TRUE
)

unlink(check_file)

if (!any(grepl("WORD_OK", word_check, fixed = TRUE))) {
  stop(
    "\nERROR: Microsoft Word could not be opened through COM automation.\n",
    "Please make sure Microsoft Word is installed."
  )
}

cat("Microsoft Word detected successfully.\n")


# ============================================================
# 5. RECURSIVE DIRECTORY SCANNER
#
# IMPORTANT:
# The structure is created recursively.
#
# For every folder:
#
#   1. Add the folder
#   2. Enter each subfolder alphabetically
#   3. Recursively process it
#   4. Return to parent
#   5. Add files belonging directly to the parent
#
# This means the hierarchy follows the REAL Windows
# parent-child structure.
# ============================================================

folder_entries <- list()

skipped_folders <- data.frame(
  folder = character(),
  error = character(),
  stringsAsFactors = FALSE
)

folders_processed <- 0
files_processed <- 0


# ------------------------------------------------------------
# Safe folder listing
# ------------------------------------------------------------

safe_list_dirs <- function(path) {
  
  result <- tryCatch(
    {
      list.dirs(
        path = path,
        full.names = TRUE,
        recursive = FALSE
      )
    },
    error = function(e) {
      
      skipped_folders <<- rbind(
        skipped_folders,
        data.frame(
          folder = path,
          error = conditionMessage(e),
          stringsAsFactors = FALSE
        )
      )
      
      character(0)
    }
  )
  
  result
}


# ------------------------------------------------------------
# Safe file listing
# ------------------------------------------------------------

safe_list_files <- function(path) {
  
  result <- tryCatch(
    {
      list.files(
        path = path,
        full.names = TRUE,
        recursive = FALSE,
        all.files = FALSE,
        include.dirs = FALSE,
        no.. = TRUE
      )
    },
    error = function(e) {
      
      skipped_folders <<- rbind(
        skipped_folders,
        data.frame(
          folder = path,
          error = conditionMessage(e),
          stringsAsFactors = FALSE
        )
      )
      
      character(0)
    }
  )
  
  result
}


# ------------------------------------------------------------
# Recursive folder processor
# ------------------------------------------------------------

process_folder <- function(path, depth) {
  
  # Count this folder
  folders_processed <<- folders_processed + 1
  
  # ----------------------------------------------------------
  # Add current folder to document structure
  # ----------------------------------------------------------
  
  folder_name <- basename(path)
  
  folder_entries[[length(folder_entries) + 1]] <<- list(
    type = "folder",
    name = folder_name,
    depth = min(depth, 9)
  )
  
  
  # ----------------------------------------------------------
  # Get immediate subfolders
  # ----------------------------------------------------------
  
  subfolders <- safe_list_dirs(path)
  
  # Remove the current directory if returned unexpectedly
  subfolders <- subfolders[
    normalizePath(
      subfolders,
      winslash = "\\",
      mustWork = FALSE
    ) !=
      normalizePath(
        path,
        winslash = "\\",
        mustWork = FALSE
      )
  ]
  
  # Sort folders alphabetically
  if (length(subfolders) > 0) {
    
    subfolder_names <- basename(subfolders)
    
    subfolders <- subfolders[
      order(
        tolower(subfolder_names),
        subfolder_names
      )
    ]
  }
  
  
  # ----------------------------------------------------------
  # RECURSIVELY PROCESS EACH SUBFOLDER
  # ----------------------------------------------------------
  
  if (length(subfolders) > 0) {
    
    for (subfolder in subfolders) {
      
      # Check that it is actually accessible
      accessible <- tryCatch(
        {
          file.access(subfolder, mode = 4) == 0
        },
        error = function(e) FALSE
      )
      
      if (!accessible) {
        
        skipped_folders <<- rbind(
          skipped_folders,
          data.frame(
            folder = subfolder,
            error = "Folder is not accessible.",
            stringsAsFactors = FALSE
          )
        )
        
        cat(
          "\nSKIPPED FOLDER:\n",
          subfolder,
          "\nReason: Folder is not accessible.\n",
          sep = ""
        )
        
        next
      }
      
      # Recursively process folder
      process_folder(
        subfolder,
        depth + 1
      )
    }
  }
  
  
  # ----------------------------------------------------------
  # NOW PROCESS FILES BELONGING DIRECTLY TO THIS FOLDER
  #
  # Files are deliberately added AFTER all subfolders.
  # ----------------------------------------------------------
  
  files <- safe_list_files(path)
  
  if (length(files) > 0) {
    
    # Keep only actual files
    files <- files[
      file.info(files)$isdir %in% FALSE
    ]
    
    # Sort alphabetically
    file_names <- basename(files)
    
    files <- files[
      order(
        tolower(file_names),
        file_names
      )
    ]
    
    # Add files
    for (file in files) {
      
      folder_entries[[length(folder_entries) + 1]] <<- list(
        type = "file",
        name = basename(file),
        depth = depth
      )
      
      files_processed <<- files_processed + 1
      
      # Progress display every 100 files
      if (files_processed %% 100 == 0) {
        
        cat(
          "Files processed:",
          files_processed,
          "\n"
        )
      }
    }
  }
}


# ============================================================
# 6. START SCANNING
# ============================================================

cat("\n")
cat("============================================================\n")
cat("SCANNING MODIOLUS PROJECT FOLDER\n")
cat("============================================================\n")
cat("\n")

cat("Project folder:\n")
cat(project_folder, "\n\n")

cat("Starting recursive scan...\n\n")


process_folder(
  project_folder,
  depth = 1
)


# ============================================================
# 7. DISPLAY SCAN SUMMARY
# ============================================================

cat("\n")
cat("============================================================\n")
cat("SCAN COMPLETE\n")
cat("============================================================\n")

cat("Folders processed:", folders_processed, "\n")
cat("Files processed:", files_processed, "\n")
cat("Skipped folders:", nrow(skipped_folders), "\n")


# ============================================================
# 8. CREATE A TEMPORARY DATA FILE FOR POWERSHELL
#
# R writes the hierarchy to a temporary tab-delimited file.
#
# PowerShell then reads it and creates the Word document.
# ============================================================

entries_file <- tempfile(
  pattern = "modiolus_folder_structure_",
  fileext = ".txt"
)


# ------------------------------------------------------------
# Convert structure to a simple tab-separated format
#
# type    depth    name
# folder  1        Modiolus_modiolus_PaulRuben
# folder  2        Project_Modiolus_modiolus_IRELAND
# file    2        example.xlsx
# ------------------------------------------------------------

entry_lines <- vapply(
  folder_entries,
  function(x) {
    
    # Replace tabs/newlines in names so the temporary
    # transfer file remains valid.
    safe_name <- gsub(
      "[\t\r\n]",
      " ",
      x$name
    )
    
    paste(
      x$type,
      x$depth,
      safe_name,
      sep = "\t"
    )
  },
  character(1)
)

writeLines(
  entry_lines,
  entries_file,
  useBytes = TRUE
)


# ============================================================
# 9. CREATE POWERSHELL SCRIPT
#
# This script:
#
#   - opens Microsoft Word
#   - creates a new document
#   - adds introduction
#   - applies native Heading 1-9 styles
#   - adds files using Normal style
#   - saves as DOCX
# ============================================================

ps_script <- '
param(
    [string]$EntriesFile,
    [string]$OutputFile,
    [string]$ProjectFolder,
    [string]$GeneratedDate
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# Microsoft Word constants
# ------------------------------------------------------------

$wdCollapseEnd = 0
$wdFormatDocumentDefault = 16

# ------------------------------------------------------------
# Start Word
# ------------------------------------------------------------

$word = New-Object -ComObject Word.Application

$word.Visible = $false

$document = $word.Documents.Add()

try {

    # --------------------------------------------------------
    # PAGE SETUP
    # --------------------------------------------------------

    $selection = $word.Selection

    # --------------------------------------------------------
    # TITLE
    #
    # The title itself is deliberately Normal style.
    # The folder hierarchy begins with Heading 1.
    # --------------------------------------------------------

    $selection.Style = $document.Styles.Item("Title")
    $selection.TypeText("MODIOLUS MODIOLUS")
    $selection.TypeParagraph()

    $selection.Style = $document.Styles.Item("Subtitle")
    $selection.TypeText("PROJECT FOLDER NAVIGATION")
    $selection.TypeParagraph()
    $selection.TypeParagraph()


    # --------------------------------------------------------
    # PROJECT INFORMATION
    # --------------------------------------------------------

    $selection.Style = $document.Styles.Item("Normal")

    $selection.TypeText("Root folder:")
    $selection.TypeParagraph()

    $selection.TypeText($ProjectFolder)
    $selection.TypeParagraph()

    $selection.TypeText("Generated:")
    $selection.TypeParagraph()

    $selection.TypeText($GeneratedDate)
    $selection.TypeParagraph()
    $selection.TypeParagraph()


    # --------------------------------------------------------
    # HOW TO NAVIGATE
    # --------------------------------------------------------

    $selection.Style = $document.Styles.Item("Heading 1")
    $selection.TypeText("HOW TO NAVIGATE")
    $selection.TypeParagraph()

    $selection.Style = $document.Styles.Item("Normal")

    $selection.TypeText(
        "The folder structure below follows the actual parent-child structure of the project directory."
    )

    $selection.TypeParagraph()

    $selection.TypeText(
        "Use View → Navigation Pane in Microsoft Word to browse the hierarchy."
    )

    $selection.TypeParagraph()
    $selection.TypeParagraph()


    # --------------------------------------------------------
    # READ THE R FOLDER STRUCTURE
    # --------------------------------------------------------

    $lines = Get-Content -LiteralPath $EntriesFile


    # --------------------------------------------------------
    # PROCESS EACH ENTRY IN EXACT RECURSIVE ORDER
    # --------------------------------------------------------

    foreach ($line in $lines) {

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $parts = $line -split "`t", 3

        if ($parts.Count -lt 3) {
            continue
        }

        $type = $parts[0]
        $depth = [int]$parts[1]
        $name = $parts[2]


        # ----------------------------------------------------
        # FOLDER
        # ----------------------------------------------------

        if ($type -eq "folder") {

            # Heading 1 through Heading 9.
            # Anything deeper than 9 remains Heading 9.

            if ($depth -lt 1) {
                $depth = 1
            }

            if ($depth -gt 9) {
                $depth = 9
            }

            $styleName = "Heading " + $depth

            $selection.Style = $document.Styles.Item($styleName)

            $selection.TypeText("FOLDER: " + $name)

            $selection.TypeParagraph()
        }


        # ----------------------------------------------------
        # FILE
        # ----------------------------------------------------

        elseif ($type -eq "file") {

            # Files use Normal style.
            $selection.Style = $document.Styles.Item("Normal")

            $selection.TypeText("FILE: " + $name)

            $selection.TypeParagraph()
        }
    }


    # --------------------------------------------------------
    # SAVE DOCUMENT
    # --------------------------------------------------------

    $document.SaveAs2(
        $OutputFile,
        $wdFormatDocumentDefault
    )


    # --------------------------------------------------------
    # CLOSE DOCUMENT
    # --------------------------------------------------------

    $document.Close()

    # Release document COM object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    $document = $null


    # --------------------------------------------------------
    # QUIT WORD
    # --------------------------------------------------------

    $word.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    $word = $null


    # --------------------------------------------------------
    # SUCCESS MESSAGE FOR R
    # --------------------------------------------------------

    Write-Output "WORD_DOCUMENT_CREATED"


}
catch {

    # --------------------------------------------------------
    # ERROR CLEANUP
    # --------------------------------------------------------

    if ($document -ne $null) {

        try {
            $document.Close($false)
        }
        catch {}

        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        }
        catch {}
    }

    if ($word -ne $null) {

        try {
            $word.Quit()
        }
        catch {}

        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        catch {}
    }

    Write-Error $_
    exit 1
}
'


ps_file <- tempfile(
  pattern = "create_modolius_word_",
  fileext = ".ps1"
)

writeLines(
  ps_script,
  ps_file,
  useBytes = TRUE
)


# ============================================================
# 10. CURRENT DATE AND TIME
# ============================================================

generated_date <- format(
  Sys.time(),
  "%Y-%m-%d %H:%M:%S"
)


# ============================================================
# 11. RUN POWERSHELL / MICROSOFT WORD
# ============================================================

cat("\n")
cat("============================================================\n")
cat("CREATING MICROSOFT WORD DOCUMENT\n")
cat("============================================================\n")
cat("\n")

cat("Output:\n")
cat(output_file, "\n\n")


ps_args <- c(
  "-NoProfile",
  "-ExecutionPolicy", "Bypass",
  "-File", shQuote(ps_file),
  "-EntriesFile", shQuote(entries_file),
  "-OutputFile", shQuote(output_file),
  "-ProjectFolder", shQuote(project_folder),
  "-GeneratedDate", shQuote(generated_date)
)


powershell_output <- system2(
  powershell,
  args = ps_args,
  stdout = TRUE,
  stderr = TRUE
)


# ============================================================
# 12. CLEAN TEMPORARY FILES
# ============================================================

unlink(
  c(
    entries_file,
    ps_file
  ),
  force = TRUE
)


# ============================================================
# 13. CHECK WHETHER WORD DOCUMENT WAS CREATED
# ============================================================

if (!file.exists(output_file)) {
  
  cat("\n")
  cat("============================================================\n")
  cat("ERROR\n")
  cat("============================================================\n")
  
  cat(
    "Microsoft Word did not create the expected document.\n\n"
  )
  
  if (length(powershell_output) > 0) {
    
    cat(
      paste(
        powershell_output,
        collapse = "\n"
      )
    )
    
    cat("\n")
  }
  
  stop(
    "\nThe Word document could not be created."
  )
}


# ============================================================
# 14. FINAL SUCCESS MESSAGE
# ============================================================

cat("\n")
cat("============================================================\n")
cat("SUCCESS - WORD DOCUMENT CREATED\n")
cat("============================================================\n")

cat("\n")

cat(
  "Document created:\n",
  output_file,
  "\n\n",
  sep = ""
)

cat(
  "Folders processed:",
  folders_processed,
  "\n"
)

cat(
  "Files processed:",
  files_processed,
  "\n"
)

cat(
  "Skipped folders:",
  nrow(skipped_folders),
  "\n"
)


# ============================================================
# 15. REPORT SKIPPED FOLDERS
# ============================================================

if (nrow(skipped_folders) > 0) {
  
  cat("\n")
  cat("============================================================\n")
  cat("SKIPPED FOLDERS\n")
  cat("============================================================\n")
  
  for (i in seq_len(nrow(skipped_folders))) {
    
    cat("\nFolder:\n")
    cat(
      skipped_folders$folder[i],
      "\n"
    )
    
    cat("Error:\n")
    cat(
      skipped_folders$error[i],
      "\n"
    )
  }
}


# ============================================================
# 16. FINAL INSTRUCTIONS
# ============================================================

cat("\n")
cat("============================================================\n")
cat("NEXT STEP\n")
cat("============================================================\n")

cat(
  "\nOpen the document in Microsoft Word and select:\n\n"
)

cat(
  "View → Navigation Pane\n\n"
)

cat(
  "The folder hierarchy will appear in Word's Navigation Pane.\n"
)

cat("\nDone.\n")

