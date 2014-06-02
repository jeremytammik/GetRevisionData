#region Namespaces
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
#endregion

namespace GetRevisionData
{
  /// <summary>
  /// External command demonstrating how to use the 
  /// Revit 2015 Revision API to retrieve and display
  /// all infoormation shown in the Revit 
  /// 'Sheet Issues/Revisions' dialogue.
  /// </summary>
  [Transaction( TransactionMode.ReadOnly )]
  public class Command : IExternalCommand
  {
    #region Unused parameter value to display string converter
    /// <summary>
    /// Extract the parameter information.
    /// By Dan Tartaglia.
    /// </summary>
    public string GetParameterInformation(
      Parameter para,
      Document doc )
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

          defName = ( id.IntegerValue >= 0 )
            ? doc.GetElement( id ).Name
            : id.IntegerValue.ToString();

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
    #endregion // Unused parameter value to display string converter

    #region Dan sample code for Revit 2015
    void f( TextWriter tw, ViewSheet oViewSheet )
    {
      Document doc = oViewSheet.Document;

      IList<ElementId> oElemIDs = oViewSheet.GetAllRevisionIds();

      if( oElemIDs.Count == 0 )
        return;

      foreach( ElementId elemID in oElemIDs )
      {
        Element oEl = doc.GetElement( elemID );

        Revision oRev = oEl as Revision;

        // Add text line to text file
        tw.WriteLine( "Rev Category Name: " + oRev.Category.Name );

        // Add text line to text file
        tw.WriteLine( "Rev Description: " + oRev.Description );

        // Add text line to text file
        tw.WriteLine( "Rev Issued: " + oRev.Issued.ToString() );

        // Add text line to text file
        tw.WriteLine( "Rev Issued By: " + oRev.IssuedBy.ToString() );

        // Add text line to text file
        tw.WriteLine( "Rev Issued To: " + oRev.IssuedTo.ToString() );

        // Add text line to text file
        tw.WriteLine( "Rev Number Type: " + oRev.NumberType.ToString() );

        // Add text line to text file
        tw.WriteLine( "Rev Date: " + oRev.RevisionDate );

        // Add text line to text file
        tw.WriteLine( "Rev Visibility: " + oRev.Visibility.ToString() );

        // Add text line to text file
        tw.WriteLine( "Rev Sequence Number: " + oRev.SequenceNumber.ToString() );

        // Add text line to text file
        tw.WriteLine( "Rev Number: " + oRev.RevisionNumber );
      }
    }
    #endregion // Dan sample code for Revit 2015

    /// <summary>
    /// A container for the revision data displayed in
    /// the Revit 'Sheet Issues/Revisions' dialogue.
    /// </summary>
    class RevisionData
    {
      public int Sequence { get; set; }
      public RevisionNumberType Numbering { get; set; }
      public string Date { get; set; }
      public string Description { get; set; }
      public bool Issued { get; set; }
      public string IssuedTo { get; set; }
      public string IssuedBy { get; set; }
      public RevisionVisibility Show { get; set; }

      public RevisionData( Revision r )
      {
        Sequence = r.SequenceNumber;
        Numbering = r.NumberType;
        Date = r.RevisionDate;
        Description = r.Description;
        Issued = r.Issued;
        IssuedTo = r.IssuedTo;
        IssuedBy = r.IssuedBy;
        Show = r.Visibility;
      }
    }

    /// <summary>
    /// Generate a Windows modeless form on the fly 
    /// and display the revision data in it in a 
    /// DataGridView.
    /// </summary>
    void DisplayRevisionData(
      List<RevisionData> revision_data,
      IWin32Window owner )
    {
      System.Windows.Forms.Form form
        = new System.Windows.Forms.Form();

      form.Size = new Size( 680, 180 );
      form.Text = "Revision Data";

      DataGridView dg = new DataGridView();
      dg.DataSource = revision_data;
      dg.AllowUserToAddRows = false;
      dg.AllowUserToDeleteRows = false;
      dg.AllowUserToOrderColumns = true;
      dg.Dock = System.Windows.Forms.DockStyle.Fill;
      dg.Location = new System.Drawing.Point( 0, 0 );
      dg.ReadOnly = true;
      dg.TabIndex = 0;
      dg.Parent = form;
      dg.AutoSize = true;
      dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
      dg.AutoResizeColumns( DataGridViewAutoSizeColumnsMode.AllCells );
      dg.AutoResizeColumns();

      form.ShowDialog( owner );
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      IWin32Window revit_window
        = new JtWindowHandle(
          ComponentManager.ApplicationWindow );

      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = uidoc.Document;

      if( doc.IsFamilyDocument )
      {
        TaskDialog.Show( "Not a Revit RVT Project",
          "This command requires an active Revit RVT file." );

        return Result.Failed;
      }

      if( !( doc.ActiveView is ViewSheet ) )
      {
        TaskDialog.Show( "Current View is not a Sheet",
          "This command requires an active sheet view." );
        return Result.Failed;
      }

      IList<ElementId> ids
        = Revision.GetAllRevisionIds( doc );

      int n = ids.Count;

      List<RevisionData> revision_data
        = new List<RevisionData>( n );

      foreach( ElementId id in ids )
      {
        Revision r = doc.GetElement( id ) as Revision;

        revision_data.Add( new RevisionData( r ) );
      }

      DisplayRevisionData( revision_data,
        revit_window );

      return Result.Succeeded;
    }
  }
}
