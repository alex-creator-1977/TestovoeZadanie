﻿Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' Allgemeine Informationen über eine Assembly werden über folgende 
' Attribute gesteuert. Ändern Sie diese Attributwerte, um die Informationen zu ändern,
' die einer Assembly zugeordnet sind.

' Werte der Assemblyattribute überprüfen

<Assembly: AssemblyTitle("ProjectAddIn2")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("ProjectAddIn2")> 
<Assembly: AssemblyCopyright("Copyright ©  2024")> 
<Assembly: AssemblyTrademark("")> 

' Durch Festlegen von "ComVisible" auf "false" werden die Typen in dieser Assembly unsichtbar 
' für COM-Komponenten. Wenn Sie auf einen Typ in dieser Assembly von 
' COM aus zugreifen müssen, legen Sie das ComVisible-Attribut für diesen Typ auf "true" fest.
<Assembly: ComVisible(False)>

'Die folgende GUID bestimmt die ID der Typbibliothek, wenn dieses Projekt für COM verfügbar gemacht wird.
<Assembly: Guid("c2291ee0-43b5-4873-a335-be42abe26d30")> 

' Versionsinformationen für eine Assembly bestehen aus den folgenden vier Werten:
'
'      Hauptversion
'      Nebenversion 
'      Buildnummer
'      Revision
'
' Sie können alle Werte angeben oder die standardmäßigen Build- und Revisionsnummern 
' übernehmen, indem Sie "*" eingeben:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
