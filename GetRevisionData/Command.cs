using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace NBBJ.NBBJProjectRevisionsTest.CS
{
  [Regeneration( RegenerationOption.Manual )]
  [Transaction( TransactionMode.Manual )]

  public class Command : IExternalCommand
  {
    // the active Revit application
    public static UIApplication m_app;

    /*------------------------------------------------------------------------------------**/
    /// <author>Dan.Tartaglia </author>                              <date>03/2010</date>
    /*--------------+---------------+---------------+---------------+---------------+------*/
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      // Get the application
      m_app = commandData.Application;

      // Get the active document
      Document doc = m_app.ActiveUIDocument.Document;

      // Display a message and exit if the active document is a family
      if( doc.IsFamilyDocument )
      {
        MessageBox.Show( "This utility should only be utilized when a Revit RVT file is active",
            "A Revit RVT file is not active.", MessageBoxButtons.OK, MessageBoxIcon.Information );
        return Result.Failed;
      }

      // Display a message and exit if the active view is not a sheet
      if( !( doc.ActiveView is ViewSheet ) )
      {
        MessageBox.Show( "This utility should only be utilized in a sheet view",
            "Active view is not a sheet.", MessageBoxButtons.OK, MessageBoxIcon.Information );
        return Result.Failed;
      }

      ProcessStart( doc );

      return Result.Succeeded;
    }

    public void ProcessStart( Document doc )
    {
      TextWriter tw = null;

      try
      {
        ViewSheet oViewSheet = doc.ActiveView
          as ViewSheet;

        if( oViewSheet == null )
          return;

        // Create a writer and open the file

        tw = new StreamWriter(
          "C:/tmp/RevisionTest.txt" );

        ParamsFromGetAllRevElements(
          oViewSheet, tw, doc );

        // Close the text file stream

        tw.Close();
      }
      catch
      {
        // Close the text file stream

        tw.Close();

        return;
      }
    }

    public void ParamsFromGetAllRevElements(
      ViewSheet oViewSheet,
      TextWriter tw,
      Document doc )
    {
      try
      {
        IList<ElementId> oElemIDs
          = oViewSheet.GetAllProjectRevisionIds();

        if( oElemIDs.Count == 0 )
          return;

        foreach( ElementId elemID in oElemIDs )
        {
          Element oEl = doc.GetElement( elemID );

          bool blnHidden = oEl.IsHidden( oViewSheet );

          // Add text line to text file
          tw.WriteLine( "Hidden = "
            + blnHidden.ToString() );

          // Add text line to text file
          tw.WriteLine( "Element Name: "
            + oEl.Name );

          Parameter oParamRevEnum = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_ENUMERATION );

          if( oParamRevEnum != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevEnum = "
              + GetParameterInformation(
                oParamRevEnum, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevEnum = null" );

          Parameter oParamRevDate = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_DATE );

          if( oParamRevDate != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevDate = "
              + GetParameterInformation(
                oParamRevDate, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevDate = null" );

          Parameter oParamRevDescrip = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_DESCRIPTION );

          if( oParamRevDescrip != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevDescrip = "
              + GetParameterInformation(
                oParamRevDescrip, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevDescrip = null" );

          Parameter oParamRevIssued = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED );

          if( oParamRevIssued != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevIssued = "
              + GetParameterInformation(
                oParamRevIssued, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevIssued = null" );

          Parameter oParamRevIssuedBy = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED_BY );

          if( oParamRevIssuedBy != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevIssuedBy = "
              + GetParameterInformation(
                oParamRevIssuedBy, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevIssuedBy = null" );

          Parameter oParamRevIssuedTo = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED_TO );

          if( oParamRevIssuedTo != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevIssuedTo = "
              + GetParameterInformation(
                oParamRevIssuedTo, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevIssuedTo = null" );

          Parameter oParamRevNumber = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_REVISION_NUM );

          if( oParamRevNumber != null )
            // Add text line to text file
            tw.WriteLine( "oParamRevNumber = "
              + GetParameterInformation(
                oParamRevNumber, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamRevNumber = null" );

          Parameter oParamSeqNumber = oEl.get_Parameter(
            BuiltInParameter.PROJECT_REVISION_SEQUENCE_NUM );

          if( oParamSeqNumber != null )
            // Add text line to text file
            tw.WriteLine( "oParamSeqNumber = "
              + GetParameterInformation(
                oParamSeqNumber, doc ) );
          else
            // Add text line to text file
            tw.WriteLine( "oParamSeqNumber = null" );

          // Add text line to text file
          tw.WriteLine( "==============================" );
        }
      }
      catch
      {
      }
    }

    /*------------------------------------------------------------------------------------**/
    /// <summary>
    /// Extract the parameter information 
    /// </summary>
    /// <returns> string </returns>
    /// <author>Dan.Tartaglia </author>                              <date>04/2010</date>
    /*--------------+---------------+---------------+---------------+---------------+------*/
    public string GetParameterInformation(
      Parameter para,
      Document document )
    {
      string defName = "";

      // Use different method to get parameter 
      // data according to the storage type

      switch( para.StorageType )
      {
        // Determine the parameter type 

        case StorageType.Double:

          // Convert the number into Metric

          defName = para.AsValueString();
          break;

        case StorageType.ElementId:

          // Find out the name of the element

          Autodesk.Revit.DB.ElementId id
            = para.AsElementId();

          if( id.IntegerValue >= 0 )
          {
            defName = document.GetElement( id ).Name;
          }
          else
          {
            defName = id.IntegerValue.ToString();
          }
          break;

        case StorageType.Integer:
          if( ParameterType.YesNo
            == para.Definition.ParameterType )
          {
            if( para.AsInteger() == 0 )
            {
              defName = "False";
            }
            else
            {
              defName = "True";
            }
          }
          else
          {
            defName = para.AsInteger().ToString();
          }
          break;

        case StorageType.String:
          defName = para.AsString();
          break;

        default:
          defName = "Unexposed parameter";
          break;
      }
      return defName;
    }
  }
}
