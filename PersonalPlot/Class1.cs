using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Geometry;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Collections.Specialized;
using System.Drawing.Printing;
using System.IO;
using WinForms = System.Windows.Forms;

namespace PersonalPlot
{

    //Jig to be let user move a window over print area
    class BlockJig : EntityJig
    {
        Point3d insertPoint, mActualPoint;
        double _angle;

        public BlockJig(BlockReference br)
            : base(br)
        {
            insertPoint = br.Position;
            _angle = 0;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions jigOpts = new JigPromptPointOptions();
            jigOpts.UserInputControls = (UserInputControls.Accept3dCoordinates)
                | UserInputControls.NoZeroResponseAccepted
                | UserInputControls.NoNegativeResponseAccepted;
            jigOpts.SetMessageAndKeywords("\nSurround plot area: " + "\nTab to rotate window [ROtate90]:","ROtate90");

            PromptPointResult ppr = prompts.AcquirePoint(jigOpts);
            if (ppr.Status == PromptStatus.Keyword)
            {
                if (ppr.StringResult == "ROtate90")
                {
                    //CW subtract 90 deg & normalize the angle between 0 & 360
                    _angle -= Math.PI / 2;
                    while (_angle < Math.PI * 2)
                    { _angle += Math.PI * 2; }
                }
                return SamplerStatus.OK;
            }
            else if (ppr.Status == PromptStatus.OK)
            {
                if (mActualPoint == ppr.Value)
                    return SamplerStatus.NoChange;
                else
                    mActualPoint = ppr.Value;

                return SamplerStatus.OK;
            }

            return SamplerStatus.Cancel;
        }

        protected override bool Update()
        {
            insertPoint = mActualPoint;
            try
            {
                ((BlockReference)Entity).Position = insertPoint;
                ((BlockReference)Entity).Rotation = _angle;
            }
            catch (System.Exception)
            { return false; }

            return true;
        }

        public Entity GetEntity()
        { return Entity; }
    }


    public class Class1
    {
        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);

        [CommandMethod("PRINTSCALED")]
        static public void plotterscaletofit()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //ask user for print window
                Point3d first;
                PromptPointOptions ppo = new PromptPointOptions("\nSelect First corner of plot area: ");
                ppo.AllowNone = false;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.OK)
                { first = ppr.Value; }
                else
                    return;

                Point3d second;
                PromptCornerOptions pco = new PromptCornerOptions("\nSelect second corner of the plot area.", first);
                ppr = ed.GetCorner(pco);
                if (ppr.Status == PromptStatus.OK)
                { second = ppr.Value; }
                else
                    return;

                //convert from UCS to DCS
                Extents2d window = coordinates(first, second);

                //if the current view is paperspace then need to set up a viewport first
                if(LayoutManager.Current.CurrentLayout != "Model")
                { }

                //set up the plotter
                PlotInfo pi = plotSetUp(window, tr, db, ed, true, false);

                //call plotter engine to run
                plotEngine(pi, "Nameless", doc, ed, false);

                tr.Dispose();
            }
        }

        //create a viewport so that plotting from the paperspace is possibe
        static public void makeViewPort(Extents3d window)
        {
            Viewport acVport = new Viewport();
            acVport.Width = window.MaxPoint.X - window.MinPoint.X;
            acVport.Height = window.MaxPoint.Y - window.MinPoint.Y;
            acVport.CenterPoint = new Point3d(acVport.Width/2 + window.MinPoint.X,
                acVport.Height/2 + window.MinPoint.Y,
                0); //dist/2 +minpoint
        }

        //instead of creating select window,
        //need to pick a center point for the 8.5 x 11 true size
        //use jig to frame size and use tab for portrait vs landscape
        [CommandMethod("PRINTTRUESCALE")]
        static public void plottertruescaled()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //get a point from user
                //use a jig to create a "window of the printed area"
                //point should be the center of the window
                //"tab" rotates from portrait and landscape

                //ask user for print window
                Extents3d plotWindow = BlockInserter(db, ed, doc);
                Extents2d window = coordinates(plotWindow.MinPoint, plotWindow.MaxPoint);

                //set up the plotter
                PlotInfo pi = plotSetUp(window, tr, db, ed, false, false);

                //call plotter engine to run
                plotEngine(pi, "Nameless", doc, ed, false);

                tr.Dispose();
            }
        }

        //loads the block into the blocktable if it isnt there
        //then uses a jig to find extents of print area
        public static Extents3d BlockInserter(Database db, Editor ed, Document doc)
        {
            //create a window to show the plot area
            Polyline polyPlot = new Polyline();            
            polyPlot.AddVertexAt(0, new Point2d(-4.25, 5.5), 0, 0, 0);
            polyPlot.AddVertexAt(1, new Point2d(4.25, 5.5), 0, 0, 0);
            polyPlot.AddVertexAt(2, new Point2d(4.25, -5.5), 0, 0, 0);
            polyPlot.AddVertexAt(3, new Point2d(-4.25, -5.5), 0, 0, 0);
            polyPlot.Closed = true;

            using (Transaction trCurrent = db.TransactionManager.StartTransaction())
            {
                //open block table for read
                BlockTable btCurrent = trCurrent.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                //check if spec is already loaded into drawing
                ObjectId blkRecId = ObjectId.Null;
                if (!btCurrent.Has("PlotGuide"))
                {
                    using (BlockTableRecord newRec = new BlockTableRecord())
                    {
                        newRec.Name = "PlotGuide";
                        newRec.Origin = new Point3d(0, 0, 0);
                        //add the polyline to the block
                        newRec.AppendEntity(polyPlot);
                        //from read to write
                        btCurrent.UpgradeOpen();
                        btCurrent.Add(newRec);
                        trCurrent.AddNewlyCreatedDBObject(newRec, true);

                        blkRecId = newRec.Id;
                    }
                }
                else
                { blkRecId = btCurrent["PlotGuide"]; }

                //now insert block into current space using our jig
                Point3d stPnt = new Point3d(0, 0, 0);
                BlockReference br = new BlockReference(stPnt, blkRecId);
                BlockJig entJig = new BlockJig(br);

                //use jig
                Extents3d printWindow = new Extents3d();
                //loop as jig is running, to get keywords/key strokes for rotation
                PromptStatus stat = PromptStatus.Keyword;
                while (stat == PromptStatus.Keyword)
                {
                    //use iMsg filter
                    var filt = new TxtRotMsgFilter(doc);

                    WinForms.Application.AddMessageFilter(filt);
                    PromptResult pr = ed.Drag(entJig);
                    WinForms.Application.RemoveMessageFilter(filt);

                    stat = pr.Status;
                    if (stat == PromptStatus.OK)
                    { printWindow = entJig.GetEntity().GeometricExtents; }

                    //commit changes
                }
                trCurrent.Commit();

                return printWindow;
            }
        }

        [CommandMethod("PRINTPDF")]
        static public void plotterPDF()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //ask user for print window
                Point3d first;
                PromptPointOptions ppo = new PromptPointOptions("\nSelect First corner of plot area: ");
                ppo.AllowNone = false;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.OK)
                { first = ppr.Value; }
                else
                    return;

                Point3d second;
                PromptCornerOptions pco = new PromptCornerOptions("\nSelect second corner of the plot area.", first);
                ppr = ed.GetCorner(pco);
                if (ppr.Status == PromptStatus.OK)
                { second = ppr.Value; }
                else
                    return;

                //convert from UCS to DCS
                Extents2d window = coordinates(first, second);

                //set up the plotter
                PlotInfo pi = plotSetUp(window, tr, db, ed, true, true);

                string fileName = "PDF";
                PromptStringOptions pso = new PromptStringOptions("Enter file Name: ");
                pso.AllowSpaces = true;
                PromptResult res = ed.GetString(pso);
                if (res.Status == PromptStatus.OK)
                { fileName = res.StringResult; }
                else
                    return;                                
                    
                //call plotter engine to run
                plotEngine(pi, fileName, doc, ed, true);

                tr.Dispose();
            }
        }

        // A PlotEngine does the actual plotting
        // (can also create one for Preview)
        //***NOTE- always be sure that back ground plotting is off, in code and the users computer.
        static void plotEngine(PlotInfo pi, string name, Document doc, Editor ed, bool pdfout)
        {
            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
            {
                PlotEngine pe = PlotFactory.CreatePublishEngine();
                using (pe)
                {
                    // Create a Progress Dialog to provide info or allow the user to cancel
                    PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true);
                    using (ppd)
                    {
                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
                        ppd.LowerPlotProgressRange = 0;
                        ppd.UpperPlotProgressRange = 100;
                        ppd.PlotProgressPos = 0;

                        // Let's start the plot, at last
                        ppd.OnBeginPlot();
                        ppd.IsVisible = true;
                        pe.BeginPlot(ppd, null);

                        // We'll be plotting a single document
                        //name should be file location + prompeted answer
                        string fileLoc = Path.GetDirectoryName(doc.Name);
                        pe.BeginDocument(pi, doc.Name, null, 1, pdfout, fileLoc + @"\" + name);

                        // Which contains a single sheet
                        ppd.OnBeginSheet();
                        ppd.LowerSheetProgressRange = 0;
                        ppd.UpperSheetProgressRange = 100;
                        ppd.SheetProgressPos = 0;

                        PlotPageInfo ppi = new PlotPageInfo();
                        pe.BeginPage(ppi, pi, true, null);
                        pe.BeginGenerateGraphics(null);
                        pe.EndGenerateGraphics(null);

                        // Finish the sheet
                        pe.EndPage(null);
                        ppd.SheetProgressPos = 100;
                        ppd.OnEndSheet();

                        // Finish the document
                        pe.EndDocument(null);

                        // And finish the plot
                        ppd.PlotProgressPos = 100;
                        ppd.OnEndPlot();
                        pe.EndPlot(null);
                    }
                }
            }

            else
            {
                ed.WriteMessage("\nAnother plot is in progress.");
            }
        }

        //acquire the extents of the frame and convert them from UCS to DCS, in case of view rotation
        static public Extents2d coordinates(Point3d firstInput, Point3d secondInput)
        {
            double minX;
            double minY;
            double maxX;
            double maxY;

            //sort through the values to be sure that the correct first and second are assigned
            if (firstInput.X < secondInput.X)
            { minX = firstInput.X; maxX = secondInput.X; }
            else
            { maxX = firstInput.X; minX = secondInput.X; }

            if (firstInput.Y < secondInput.Y)
            { minY = firstInput.Y; maxY = secondInput.Y; }
            else
            { maxY = firstInput.Y; minY = secondInput.Y; }


            Point3d first = new Point3d(minX, minY, 0);
            Point3d second = new Point3d(maxX, maxY, 0);
            //converting numbers to something the system uses (DCS) instead of UCS
            ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
            double[] firres = new double[] { 0, 0, 0 };
            double[] secres = new double[] { 0, 0, 0 };
            //convert points
            acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
            acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
            Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);

            return window;
        }

        //determine if the plot is landscape or portrait based on which side is longer
        static public PlotRotation orientation(Extents2d ext)
        {
            PlotRotation portrait = PlotRotation.Degrees180;
            PlotRotation landscape = PlotRotation.Degrees270;
            double width = ext.MinPoint.X - ext.MaxPoint.X;
            double height = ext.MinPoint.Y - ext.MaxPoint.Y;
            if (Math.Abs(width) > Math.Abs(height))
            { return landscape; }
            else
            { return portrait; }
        }

        //set up plotinfo
        static public PlotInfo plotSetUp(Extents2d window, Transaction tr, Database db, Editor ed, bool scaleToFit, bool pdfout)
        {
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                // We need a PlotInfo object linked to the layout
                PlotInfo pi = new PlotInfo();
                pi.Layout = btr.LayoutId;

                //current layout
                Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

                // We need a PlotSettings object based on the layout settings which we then customize
                PlotSettings ps = new PlotSettings(lo.ModelType);
                ps.CopyFrom(lo);

                //The PlotSettingsValidator helps create a valid PlotSettings object
                PlotSettingsValidator psv = PlotSettingsValidator.Current;

                //set rotation
                psv.SetPlotRotation(ps, orientation(window)); //perhaps put orientation after window setting window??

                // We'll plot the window, centered, scaled, landscape rotation
                psv.SetPlotWindowArea(ps, window);
                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);//breaks here on some drawings                

                // Set the plot scale
                psv.SetUseStandardScale(ps, true);
                if (scaleToFit == true)
                { psv.SetStdScaleType(ps, StdScaleType.ScaleToFit); }
                else
                { psv.SetStdScaleType(ps, StdScaleType.StdScale1To1); }

                // Center the plot
                psv.SetPlotCentered(ps, true);//finding best location

                //get printerName from system settings
                PrinterSettings settings = new PrinterSettings();
                string defaultPrinterName = settings.PrinterName;

                psv.RefreshLists(ps);
                // Set Plot device & page size 
                // if PDF set it up for some PDF plotter
                if (pdfout == true)
                {
                    psv.SetPlotConfigurationName(ps, "DWG to PDF.pc3", null);
                    var mns = psv.GetCanonicalMediaNameList(ps);
                    if (mns.Contains("ANSI_expand_A_(8.50_x_11.00_Inches)"))
                    { psv.SetCanonicalMediaName(ps, "ANSI_expand_A_(8.50_x_11.00_Inches)"); }
                    else
                    { string mediaName = setClosestMediaName(psv, ps, 8.5, 11, true); }
                }
                else
                {
                    psv.SetPlotConfigurationName(ps, defaultPrinterName, null);
                    var mns = psv.GetCanonicalMediaNameList(ps);
                    if (mns.Contains("Letter"))
                    { psv.SetCanonicalMediaName(ps, "Letter"); }
                    else
                    { string mediaName = setClosestMediaName(psv, ps, 8.5, 11, true); }
                }

                //rebuilts plotter, plot style, and canonical media lists
                //(must be called before setting the plot style)
                psv.RefreshLists(ps);

                //ps.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;
                //ps.ShadePlotResLevel = ShadePlotResLevel.Normal;

                //plot options
                //ps.PrintLineweights = true;
                //ps.PlotTransparency = false;
                //ps.PlotPlotStyles = true;
                //ps.DrawViewportsFirst = true;
                //ps.CurrentStyleSheet

                // Use only on named layouts - Hide paperspace objects option
                // ps.PlotHidden = true;

                //psv.SetPlotRotation(ps, PlotRotation.Degrees180);


                //plot table needs to be the custom heavy lineweight for the Uphol specs 
                psv.SetCurrentStyleSheet(ps, "monochrome.ctb");

                // We need to link the PlotInfo to the  PlotSettings and then validate it
                pi.OverrideSettings = ps;
                PlotInfoValidator piv = new PlotInfoValidator();
                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                piv.Validate(pi);

                return pi;
            }
        }

        //if the media size doesn't exist, this will search media list for best match
        // 8.5 x 11 should be there
        private static string setClosestMediaName(PlotSettingsValidator psv, PlotSettings ps,
            double pageWidth, double pageHeight, bool matchPrintableArea)
        {
            //get all of the media listed for plotter
            StringCollection mediaList = psv.GetCanonicalMediaNameList(ps);
            double smallestOffest = 0.0;
            string selectedMedia = string.Empty;
            PlotRotation selectedRot = PlotRotation.Degrees000;

            foreach (string media in mediaList)
            {
                psv.SetCanonicalMediaName(ps, media);

                double mediaWidth = ps.PlotPaperSize.X;
                double mediaHeight = ps.PlotPaperSize.Y;

                if (matchPrintableArea)
                {
                    mediaWidth -= (ps.PlotPaperMargins.MinPoint.X + ps.PlotPaperMargins.MaxPoint.X);
                    mediaHeight -= (ps.PlotPaperMargins.MinPoint.Y + ps.PlotPaperMargins.MaxPoint.Y);
                }

                PlotRotation rot = PlotRotation.Degrees090;

                //check that we are not outside the media print area
                if (mediaWidth < pageWidth || mediaHeight < pageHeight)
                {
                    //Check if turning paper will work
                    if (mediaHeight < pageWidth || mediaWidth >= pageHeight)
                    {
                        //still too small
                        continue;
                    }
                    rot = PlotRotation.Degrees090;
                }

                double offset = Math.Abs(mediaWidth * mediaHeight - pageWidth * pageHeight);

                if (selectedMedia == string.Empty || offset < smallestOffest)
                {
                    selectedMedia = media;
                    smallestOffest = offset;
                    selectedRot = rot;

                    if (smallestOffest == 0)
                        break;
                }
            }
            psv.SetCanonicalMediaName(ps, selectedMedia);
            psv.SetPlotRotation(ps, selectedRot);
            return selectedMedia;
        }
    }

    //iMessage filter from Kean Walmsley(Through the interface) to let the jig be rotated while jigging
    public class TxtRotMsgFilter : WinForms.IMessageFilter
    {
        [DllImport( "user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]

        public static extern short GetKeyState(int keyCode);

        const int WM_KEYDOWN = 256;
        const int VK_CONTROL = 17;

        private Document _doc = null;

        public TxtRotMsgFilter(Document doc)
        { _doc = doc; }

        public bool PreFilterMessage(ref WinForms.Message m)
        {
            if (
                m.Msg == WM_KEYDOWN &&
                m.WParam == (IntPtr)WinForms.Keys.Tab &&
                GetKeyState(VK_CONTROL) >= 0)
            {
                _doc.SendStringToExecute("_RO", true, false, false);
                return true;
            }
            return false;
        }

    }
}
