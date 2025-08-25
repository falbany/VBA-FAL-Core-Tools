Attribute VB_Name = "FALCore"
' **************************************************************************************
' Module    : FALCore
' Author    : Florent ALBANY
' Website   :
' Purpose   : Central documentation and information module for the FALCore VBA Suite.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2025-07-29              Initial Release
'---------------------------------------------------------------------------------------
' @description
' This module serves as the entry point and "About" page for the FALCore suite.
' FALCore is a comprehensive library of VBA modules designed to accelerate
' application development in Microsoft Excel by providing robust, reusable,
' and well-documented functions for common tasks.
'
' The suite is organized into the following modules, each prefixed with "Fal":
'
'   - FalWork:      Functions for creating, manipulating, and managing Workbooks
'                   and Worksheets.
'
'   - FalFile:      Utilities for file and folder operations, including reading,
'                   writing, copying, moving, and zipping files.
'
'   - FalArray:     A powerful set of tools for creating, manipulating, and
'                   querying 1D, 2D, 3D, and 4D arrays.
'
'   - FalLog:       A flexible logging utility with configurable levels (Error,
'                   Warning, Info, Debug) and destinations (Immediate Window, File).
'
' Each module is designed to be self-contained where possible and follows
' consistent naming conventions.
' **************************************************************************************

Option Explicit

Public Const FALCORE_VERSION As String = "1.0.0"

Public Sub FALCore_About()
    ' @brief Displays a message box with information about the FALCore suite.
    Dim msg As String
    msg = "FALCore VBA Suite" & vbCrLf & _
          "Version: " & FALCORE_VERSION & vbCrLf & _
          "Author: Florent ALBANY" & vbCrLf & vbCrLf & _
          "A collection of robust modules designed to accelerate VBA development."
    MsgBox msg, vbOKOnly + vbInformation, "About FALCore Suite"
End Sub

