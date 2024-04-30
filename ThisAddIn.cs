using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace CTI_Utilities
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }

    public class AddAdvise
    {

        /// <summary>The EventSink class will sink the events added in this
        /// class. It will write event information to the debug output window.
        /// </summary>
        private EventSink eventHandler;

        /// <summary>This method is the class constructor.</summary>
        public AddAdvise()
        {

            // No initialization is required.
        }

        /// <summary>This method uses the AddAdvise method on an event list to
        /// tell Visio which events to monitor and which object will handle
        /// those events. This method also shows how to use the AddAdvise method
        /// at both the application and document levels, each of which has its
        /// own EventList collection.</summary>
        /// <param name="theApplication">Reference to the Visio application
        /// object</param>
        public void DemoAddAdvise(
            Microsoft.Office.Interop.Visio.Application theApplication)
        {

            // Declare visEvtAdd as a 2-byte value to avoid a run-time overflow
            // error.
            const short visEvtAdd = -32768;

            Microsoft.Office.Interop.Visio.EventList eventsDocument;
            Microsoft.Office.Interop.Visio.EventList eventsApplication;
            Microsoft.Office.Interop.Visio.Document addedDocument;

            if (theApplication == null)
            {
                return;
            }

            try
            {

                eventHandler = new EventSink();

                // Create a new drawing.
                addedDocument = Globals.ThisAddIn.Application.ActiveDocument;

                // Get the EventList collection of this Application object.
                eventsApplication = theApplication.EventList;

                // Get the EventList collection of this document.
                eventsDocument = addedDocument.EventList;

                // Add events for which Visio will send notification to
                // the EventSink class.

                // Add the QueryCancelSelectionDelete event.
                eventsDocument.AddAdvise((short)Microsoft.Office.Interop.Visio.
                    VisEventCodes.visEvtCodeQueryCancelSelDel,
                    eventHandler, "", "");

                //Add the BeforeShapeDelete event.
                eventsDocument.AddAdvise(((short)Microsoft.Office.Interop.Visio.
                     VisEventCodes.visEvtDel + (short)Microsoft.Office.Interop.
                     Visio.VisEventCodes.visEvtShape),
                    eventHandler, "", "");

                // Add the PageAdded event.
                eventsDocument.AddAdvise((short)Microsoft.Office.Interop.Visio.
                     VisEventCodes.visEvtPage + visEvtAdd,
                    eventHandler, "", "");

                // Add the ShapeAdded event.
                eventsDocument.AddAdvise((short)Microsoft.Office.Interop.Visio.
                    VisEventCodes.visEvtShape + visEvtAdd,
                    eventHandler, "", "");

                // Add the BeforeDocumentClose event.
                eventsDocument.AddAdvise(((short)Microsoft.Office.Interop.Visio.
                     VisEventCodes.visEvtDel + (short)Microsoft.Office.Interop.
                     Visio.VisEventCodes.visEvtDoc),
                    eventHandler, "", "");

                // Add the ApplicationQuit event.
                eventsApplication.AddAdvise(((short)Microsoft.Office.Interop.
                     Visio.VisEventCodes.visEvtApp + (short)Microsoft.Office.
                     Interop.Visio.VisEventCodes.visEvtBeforeQuit),
                    eventHandler, "", "");

                // Add the WindowTurnToPage event.
                eventsApplication.AddAdvise(((short)Microsoft.Office.Interop.
                     Visio.VisEventCodes.visEvtCodeWinPageTurn),
                    eventHandler, "", "");
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);

            }
        }
    }

    public class EventSink
        : Microsoft.Office.Interop.Visio.IVisEventProc
    {

        /// <summary>visEvtAdd is declared as a 2-byte value to avoid a run-time
        /// overflow error.</summary>
        private const short visEvtAdd = -32768;

        private const string eventSinkCaption = "Event Sink";
        private const string tab = "\t";
        private System.Collections.Specialized.StringDictionary
            eventDescriptions;

        /// <summary>The constructor initializes the event descriptions
        /// dictionary.</summary>
        public EventSink()
        {

            initializeStrings();
        }

        /// <summary>This method is called by Visio when an event in the
        /// EventList collection has been triggered. This method is an
        /// implementation of IVisEventProc.VisEventProc method.</summary>
        /// <param name="eventCode">Event code of the event that fired</param>
        /// <param name="source">Reference to source of the event</param>
        /// <param name="eventId">Unique identifier of the event object that 
        /// raised the event</param>
        /// <param name="eventSequenceNumber">Relative position of the event in 
        /// the event list</param>
        /// <param name="subject">Reference to the subject of the event</param>
        /// <param name="moreInformation">Additional information for the event
        /// </param>
        /// <returns>False to allow a QueryCancel operation or True to cancel 
        /// a QueryCancel operation. The return value is ignored by Visio unless 
        /// the event is a QueryCancel event.</returns>
        /// <seealso cref="Microsoft.Office.Interop.Visio.IVisEventProc"></seealso>
        public object VisEventProc(
            short eventCode,
            object source,
            int eventId,
            int eventSequenceNumber,
            object subject,
            object moreInformation)
        {

            string message = "";
            string name = "";
            string eventInformation = "";
            object returnValue = true;

            if (source == null)
            {
                return null;
            }

            Microsoft.Office.Interop.Visio.Application subjectApplication = null;
            Microsoft.Office.Interop.Visio.Document subjectDocument = null;
            Microsoft.Office.Interop.Visio.Page subjectPage = null;
            Microsoft.Office.Interop.Visio.Master subjectMaster = null;
            Microsoft.Office.Interop.Visio.Selection subjectSelection = null;
            Microsoft.Office.Interop.Visio.Shape subjectShape = null;
            Microsoft.Office.Interop.Visio.Cell subjectCell = null;
            Microsoft.Office.Interop.Visio.Connects subjectConnects = null;
            Microsoft.Office.Interop.Visio.Style subjectStyle = null;
            Microsoft.Office.Interop.Visio.Window subjectWindow = null;
            Microsoft.Office.Interop.Visio.MouseEvent subjectMouseEvent = null;
            Microsoft.Office.Interop.Visio.KeyboardEvent subjectKeyboardEvent = null;
            Microsoft.Office.Interop.Visio.DataRecordset subjectDataRecordset = null;
            Microsoft.Office.Interop.Visio.DataRecordsetChangedEvent subjectDataRecordsetChangedEvent = null;
            Microsoft.Office.Interop.Visio.RelatedShapePairEvent subjectRelatedShapePairEvent = null;
            Microsoft.Office.Interop.Visio.MovedSelectionEvent subjectMovedSelectionEvent = null;
            Microsoft.Office.Interop.Visio.ValidationRuleSet subjectValdationRuleSet = null;

            try
            {

                switch (eventCode)
                {

                    // Document event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtDoc + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBefDocSave:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBefDocSaveAs:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeDocDesign:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtDoc + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtDoc + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelDocClose:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeDocCreate:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeDocOpen:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeDocSave:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeDocSaveAs:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeDocRunning:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelDocClose:

                        // Subject object is a Document
                        // Eventinfo may be non empty. 
                        //  (1) For DocumentChanged Event it may indicate what 
                        //   changed, e.g.  /pagereordered.  
                        //  (2) For the Save and SaveAs related events, the eventInfo may contain
                        //   version information with the format /version=X where X is the file 
                        //   version number.  For SaveAs events, the eventInfo also contains the 
                        //   the full path for the save as action, in the format /saveasfile=
                        //   where the full path directly follows the equal sign.  If the save  
                        //   action is the result AutoSave, then the string /saveasfile= will be 
                        //   replaced by /autosavefile=.
                        //  (3) For RemoveHiddenInformation, the eventInfo
                        //   indicates the data that were removed. The various types 
                        //   are represented by the following strings: 
                        //   /visRHIPersonalInfo, /visRHIMasters, /visRHIStyles,
                        //   /visRHIDataRecordsets, /visRHIValidationRules. The /visRHIStyles
                        //   string appears when themes, datagraphics or styles were removed.

                        subjectDocument =
                            (Microsoft.Office.Interop.Visio.Document)subject;
                        subjectApplication = subjectDocument.Application;
                        name = subjectDocument.Name;
                        break;

                    // Page event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtPage + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtPage + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtPage + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelPageDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelPageDel:

                        // Subject object is a Page
                        subjectPage = (Microsoft.Office.Interop.Visio.Page)subject;
                        subjectApplication = subjectPage.Application;
                        name = subjectPage.Name;
                        break;

                    // Master event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtMaster + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtMaster + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelMasterDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtMaster + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelMasterDel:

                        // Subject object is a Master
                        subjectMaster = (Microsoft.Office.Interop.Visio.Master)subject;
                        subjectApplication = subjectMaster.Application;
                        name = subjectMaster.Name;
                        break;

                    // Selection event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBefSelDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeSelAdded:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelSelDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelConvertToGroup:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelUngroup:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelConvertToGroup:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelSelDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelUngroup:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelSelGroup:

                        // Subject object is a Selection
                        subjectSelection =
                            (Microsoft.Office.Interop.Visio.Selection)subject;
                        subjectApplication = subjectSelection.Application;
                        break;

                    // Shape event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtShape + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeShapeBeforeTextEdit:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtShape + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtShape + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeShapeExitTextEdit:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeShapeParentChange:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeShapeDelete:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtText + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtShapeLinkAdded:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtShapeLinkDeleted:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtShapeDataGraphicChanged:

                        // Subject Object is a shape
                        // EventInfo may be non empty:
                        // (1) For ShapeChanged event, the eventInfo indicates what changed:
                        //  The possible EventInfo strings are: /name, /data1, /data2, /data3,
                        //  /uniqueid, /ink, /listorder. 
                        //  The /ink string will only appear when ink strokes are added/deleted from 
                        //  an ink shape.
                        //  The /listorder string will only appear for a list Shape when its list 
                        //  members are re-ordered.  
                        // (2) For the ShapeLinkAdded and ShapelinkDeleted events, 
                        //  the eventInfo provides the recordset ID and rowID 
                        //  participating in the link as  
                        //  /DataRecordsetID=<ID> and /DataRowID=<ID2>

                        subjectShape =
                            (Microsoft.Office.Interop.Visio.Shape)subject;
                        subjectApplication = subjectShape.Application;
                        name = subjectShape.Name;
                        break;

                    // Cell event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCell + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtFormula + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:

                        // Subject object is a Cell
                        subjectCell =
                            (Microsoft.Office.Interop.Visio.Cell)subject;
                        subjectShape = subjectCell.Shape;
                        subjectApplication = subjectCell.Application;
                        name = subjectShape.Name + "!" + subjectCell.Name;
                        break;

                    // Connects event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtConnect + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtConnect + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:

                        // Subject object is a Connects collection
                        subjectConnects =
                            (Microsoft.Office.Interop.Visio.Connects)subject;
                        subjectApplication = subjectConnects.Application;
                        break;

                    // Style event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtStyle + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtStyle + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtStyle + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelStyleDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelStyleDel:

                        // Subject object is a Style
                        subjectStyle =
                            (Microsoft.Office.Interop.Visio.Style)subject;
                        subjectApplication = subjectStyle.Application;
                        name = subjectStyle.Name;
                        break;

                    // Window event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtWindow + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBefWinPageTurn:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtWindow + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtWindow + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeWinPageTurn:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBefWinSelDel:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelWinClose:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtWinActivate:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeWinSelChange:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeViewChanged:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelWinClose:

                        // Subject object is a Window
                        subjectWindow =
                            (Microsoft.Office.Interop.Visio.Window)subject;
                        subjectApplication = subjectWindow.Application;
                        name = subjectWindow.Caption;
                        break;

                    // Application event codes
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtAfterModal:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeAfterResume:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtAppActivate:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtAppDeactivate:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtObjActivate:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtObjDeactivate:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtBeforeModal:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtBeforeQuit:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBeforeSuspend:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeEnterScope:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeExitScope:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMarker:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeBefForcedFlush:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeAfterForcedFlush:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtNonePending:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeWinOnAddonKeyMSG:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelQuit:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeQueryCancelSuspend:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelQuit:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCancelSuspend:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtApp + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtIdle:

                        // Subject object is an Application
                        // EventInfo is empty for most of these events.  However for
                        // the Marker event, the EnterScope event and the ExitScope 
                        // event eventinfo contains the context string. 
                        subjectApplication =
                            (Microsoft.Office.Interop.Visio.Application)subject;
                        break;

                    // Keyboard events
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeKeyDown:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeKeyPress:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeKeyUp:

                        // Subject object is KeyboardEvent
                        // Note, keyboard events can be canceled.
                        subjectKeyboardEvent =
                            (Microsoft.Office.Interop.Visio.KeyboardEvent)subject;
                        break;


                    // Mouse Events
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeMouseDown:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeMouseMove:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeMouseUp:

                        // Subject object is MouseEvent Object. 
                        // Eventinfo may be non-empty for mouse move events.
                        // In that cases it indicates the drag state which is 
                        // also exposed in the DragState property of the
                        // MouseEvent object. 
                        // Note, mouse events can be canceled. 
                        subjectMouseEvent =
                            (Microsoft.Office.Interop.Visio.MouseEvent)subject;
                        break;


                    // DataRecordset events
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtDataRecordset + visEvtAdd:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtDataRecordset + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtDel:

                        // Subject object is DataRecordset
                        subjectDataRecordset =
                            (Microsoft.Office.Interop.Visio.DataRecordset)subject;
                        break;

                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtDataRecordset + (short)Microsoft.Office.Interop.Visio.
                        VisEventCodes.visEvtMod:

                        // Subject object is DataRecordsetChangedEvent object
                        subjectDataRecordsetChangedEvent =
                            (Microsoft.Office.Interop.Visio.DataRecordsetChangedEvent)subject;
                        break;

                    // Relationship Changed events for callouts and containers
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCalloutRelationshipAdded:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeCalloutRelationshipDeleted:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeContainerRelationshipAdded:
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtCodeContainerRelationshipDeleted:

                        // Subject is RelatedShapePairEvent Object
                        // For the Callout Events, the FromShapeID is the ID of callout shape and 
                        // the ToShapeID is the ID of the target shape.
                        // For the Container Events, the FromShapeID is the ID of the container shape 
                        // and the ToShapeID is the ID of the member Shape.
                        subjectRelatedShapePairEvent =
                            (Microsoft.Office.Interop.Visio.RelatedShapePairEvent)subject;
                        break;

                    // Subprocess event 
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.visEvtCodeSelectionMovedToSubprocess:

                        // Subject is MovedSelectionEvent
                        subjectMovedSelectionEvent =
                            (Microsoft.Office.Interop.Visio.MovedSelectionEvent)subject;
                        break;

                    // Validation event
                    case (short)Microsoft.Office.Interop.Visio.VisEventCodes.visEvtCodeRuleSetValidated:

                        //subject is ValidationRuleSEt
                        subjectValdationRuleSet =
                            (Microsoft.Office.Interop.Visio.ValidationRuleSet)subject;
                        break;

                    default:
                        name = "Unknown";
                        break;
                }

                // get a description for this event code
                message = getEventDescription(eventCode);

                // append the name of the subject object
                if (name.Length > 0)
                {

                    message += ": " + name;
                }

                // append event info when it is available
                if (subjectApplication != null)
                {

                    eventInformation = subjectApplication.get_EventInfo(
                        (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                        visEvtIdMostRecent);

                    if (eventInformation != null)
                    {

                        message += tab + eventInformation;
                    }
                }

                // append moreInformation when it is available
                if (moreInformation != null)
                {

                    message += tab + moreInformation.ToString();
                }

                // Get the targetArgs string from the event object. The targetArgs
                // are added to the event object in the AddAdvise method
                Microsoft.Office.Interop.Visio.EventList events = null;
                Microsoft.Office.Interop.Visio.Event thisEvent = null;
                string sourceType;
                string targetArgs = "";

                sourceType = source.GetType().FullName;
                if (sourceType ==
                    "Microsoft.Office.Interop.Visio.ApplicationClass")
                {

                    events = ((Microsoft.Office.Interop.Visio.Application)source)
                        .EventList;
                }
                else if (sourceType ==
                    "Microsoft.Office.Interop.Visio.DocumentClass")
                {

                    events = ((Microsoft.Office.Interop.Visio.Document)source)
                        .EventList;
                }
                else if (sourceType ==
                    "Microsoft.Office.Interop.Visio.PageClass")
                {

                    events = ((Microsoft.Office.Interop.Visio.Page)source)
                        .EventList;
                }

                if (events != null)
                {

                    thisEvent = events.get_ItemFromID(eventId);
                    targetArgs = thisEvent.TargetArgs;

                    // append targetArgs when it is available
                    if (targetArgs.Length > 0)
                    {

                        message += " " + targetArgs;
                    }
                }

                // Write the event info to the output window
                System.Diagnostics.Debug.WriteLine(message);

                // if this is a QueryCancel event then prompt the user
                returnValue = getQueryCancelResponse(eventCode, subject);
            }
            catch (Exception err)
            {
                System.Diagnostics.Debug.WriteLine(err.Message);
            }

            return returnValue;
        }

        /// <summary>
        /// This method prompts the user to continue or cancel. If the
        /// alertResponse value is set in this Visio instance then its value 
        /// will be used and the dialog will be suppressed.</summary>
        /// <param name="eventCode">Event code of the event that fired</param>
        /// <param name="subject">Reference to subject of the event</param>
        /// <returns>False to allow the QueryCancel operation or True to cancel 
        /// the QueryCancel operation.</returns>
        private static object getQueryCancelResponse(
            short eventCode,
            object subject)
        {

            const string docCloseCancelPrompt =
                "Are you sure you want to close the document?";
            const string pageDeleteCancelPrompt =
                "Are you sure you want to delete the page?";
            const string masterDeleteCancelPrompt =
                "Are you sure you want to delete the master?";
            const string ungroupCancelPrompt =
                "Are you sure you want to ungroup the selected shapes?";
            const string convertToGroupCancelPrompt =
                "Are you sure you want to convert the selected shapes to a group?";
            const string selectionDeleteCancelPrompt =
                "Are you sure you want to delete the selected shapes?";
            const string styleDeleteCancelPrompt =
                "Are you sure you want to delete the style?";
            const string windowCloseCancelPrompt =
                "Are you sure you want to close the window?";
            const string quitCancelPrompt =
                "Are you sure you want to quit Visio?";
            const string suspendCancelPrompt =
                "Are you sure you want suspend Visio?";
            const string groupCancelPrompt =
                "Are you sure you want to group the selected shapes?";

            Microsoft.Office.Interop.Visio.Application subjectApplication = null;
            Microsoft.Office.Interop.Visio.Document subjectDocument = null;
            Microsoft.Office.Interop.Visio.Page subjectPage = null;
            Microsoft.Office.Interop.Visio.Master subjectMaster = null;
            Microsoft.Office.Interop.Visio.Selection subjectSelection = null;
            Microsoft.Office.Interop.Visio.Style subjectStyle = null;
            Microsoft.Office.Interop.Visio.Window subjectWindow = null;
            string prompt = "";
            string subjectName = "";
            short alertResponse = 0;
            bool isQueryCancelEvent = true;
            object returnValue = false;

            switch (eventCode)
            {

                // Query Document Close
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes
                .visEvtCodeQueryCancelDocClose:

                    subjectDocument = ((Microsoft.Office.Interop.Visio.Document)
                        subject);
                    subjectName = subjectDocument.Name;
                    subjectApplication = subjectDocument.Application;
                    prompt = docCloseCancelPrompt + System.Environment.NewLine
                        + subjectName;
                    break;

                // Query Cancel Page Delete
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                visEvtCodeQueryCancelPageDel:

                    subjectPage = ((Microsoft.Office.Interop.Visio.Page)subject);
                    subjectName = subjectPage.NameU;
                    subjectApplication = subjectPage.Application;
                    prompt = pageDeleteCancelPrompt + System.Environment.NewLine
                        + subjectName;
                    break;

                // Query Cancel Master Delete
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelMasterDel:

                    subjectMaster = ((Microsoft.Office.Interop.Visio.Master)
                        subject);
                    subjectName = subjectMaster.NameU;
                    subjectApplication = subjectMaster.Application;
                    prompt = masterDeleteCancelPrompt +
                        System.Environment.NewLine + subjectName;
                    break;

                // Query Cancel Ungroup
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelUngroup:

                    subjectSelection = ((Microsoft.Office.Interop.Visio.
                        Selection)subject);
                    subjectApplication = subjectSelection.Application;
                    prompt = ungroupCancelPrompt;
                    break;

                // Query Cancel Convert To Group
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelConvertToGroup:

                    subjectSelection = ((Microsoft.Office.Interop.Visio.
                        Selection)subject);
                    subjectApplication = subjectSelection.Application;
                    prompt = convertToGroupCancelPrompt;
                    break;

                // Query Cancel Selection Delete
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelSelDel:

                    subjectSelection = ((Microsoft.Office.Interop.Visio.
                        Selection)subject);
                    subjectApplication = subjectSelection.Application;
                    prompt = selectionDeleteCancelPrompt;
                    break;

                // Query Cancel Style Delete
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelStyleDel:

                    subjectStyle = ((Microsoft.Office.Interop.Visio.Style)
                        subject);
                    subjectName = subjectStyle.NameU;
                    subjectApplication = subjectStyle.Application;
                    prompt = styleDeleteCancelPrompt + System.Environment.NewLine
                        + subjectName;
                    break;

                // Query Cancel Window Close
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelWinClose:

                    subjectWindow = ((Microsoft.Office.Interop.Visio.Window)
                        subject);
                    subjectName = subjectWindow.Caption;
                    subjectApplication = subjectWindow.Application;
                    prompt = windowCloseCancelPrompt +
                        System.Environment.NewLine + subjectName;
                    break;

                // Query Cancel Quit
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelQuit:

                    subjectApplication = (Microsoft.Office.Interop.Visio.
                        Application)subject;
                    prompt = quitCancelPrompt;
                    break;

                // Query Cancel Suspend
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelSuspend:

                    subjectApplication = (Microsoft.Office.Interop.Visio.
                        Application)subject;
                    prompt = suspendCancelPrompt;
                    break;

                // Query Cancel Group
                case (short)Microsoft.Office.Interop.Visio.VisEventCodes.
                    visEvtCodeQueryCancelSelGroup:

                    subjectSelection = ((Microsoft.Office.Interop.Visio.
                        Selection)subject);
                    subjectApplication = subjectSelection.Application;
                    prompt = groupCancelPrompt;
                    break;

                default:
                    // This event is not cancelable.
                    isQueryCancelEvent = false;
                    break;
            }

            if (isQueryCancelEvent == true)
            {

                // check for an alertResponse setting in Visio
                if (subjectApplication != null)
                {
                    alertResponse = subjectApplication.AlertResponse;
                }

                if (alertResponse != 0)
                {

                    // if alertResponse is No or Cancel then cancel this event
                    // by returning true
                    if ((alertResponse == (int)System.Windows.Forms.
                        DialogResult.No) ||
                        (alertResponse == (int)System.Windows.Forms.
                        DialogResult.Cancel))
                    {
                        returnValue = true;
                    }
                }
                else
                {

                    // alertResponse is not set so prompt the user
                    System.Windows.Forms.DialogResult result;
                    result = System.Windows.Forms.MessageBox.Show(prompt,
                        eventSinkCaption,
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.No)
                    {
                        returnValue = true;
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// This method adds an event description to the eventDescriptions 
        /// dictionary.</summary>
        /// <param name="eventCode">Event code of the event</param>
        /// <param name="description">Short description of the event</param>
        private void addEventDescription(
            short eventCode,
            string description)
        {

            string key = Convert.ToString(eventCode,
                System.Globalization.CultureInfo.InvariantCulture);
            eventDescriptions.Add(key, description);
        }

        /// <summary>
        /// This method returns a short description for the given eventCode.
        /// </summary>
        /// <param name="eventCode">Event code</param>
        /// <returns>Short description of the eventCode</returns>
        private string getEventDescription(short eventCode)
        {

            string description;
            string key;

            key = Convert.ToString(eventCode,
                System.Globalization.CultureInfo.InvariantCulture);
            description = eventDescriptions[key];

            if (description == null)
            {
                description = "NoEventDescription";
            }

            return description;
        }

        /// <summary>
        /// This method populates the eventDescriptions dictionary with a short 
        /// // description of each Visio event code.</summary>
        private void initializeStrings()
        {

            eventDescriptions =
                new System.Collections.Specialized.StringDictionary();

            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtAfterModal, "AfterModal");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeAfterResume, "AfterResume");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtAppActivate, "AppActivated");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtAppDeactivate, "AppDeactivated");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtObjActivate, "AppObjActivated");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtObjDeactivate, "AppObjDeactivated");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDoc + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDel, "BeforeDocumentClose");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBefDocSave, "BeforeDocumentSave");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBefDocSaveAs, "BeforeDocumentSaveAs");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMaster + (short)Microsoft.Office.Interop
                .Visio.VisEventCodes.visEvtDel, "BeforeMasterDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtBeforeModal, "BeforeModal");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtPage + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDel, "BeforePageDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtBeforeQuit, "BeforeQuit");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBefSelDel, "BeforeSelectionDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtShape + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDel, "BeforeShapeDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeShapeBeforeTextEdit,
                "BeforeShapeTextEdit");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtStyle + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDel, "BeforeStyleDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBeforeSuspend, "BeforeSuspend");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtWindow + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDel, "BeforeWindowClose");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBefWinPageTurn, "BeforeWindowPageTurn");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBefWinSelDel, "BeforeWindowSelDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCell + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "CellChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtConnect + visEvtAdd, "ConnectionsAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtConnect + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDel, "ConnectionsDeleted");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelConvertToGroup,
                "ConvertToGroupCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeDocDesign, "DesignModeEntered");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDoc + visEvtAdd, "DocumentAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtDoc + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "DocumentChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelDocClose, "DocumentCloseCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeDocCreate, "DocumentCreated");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeDocOpen, "DocumentOpened");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeDocSave, "DocumentSaved");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeDocSaveAs, "DocumentSavedAs");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeEnterScope, "EnterScope");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeExitScope, "ExitScope");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtFormula + (short)Microsoft.Office.Interop
                .Visio.VisEventCodes.visEvtMod, "FormulaChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeKeyDown, "KeyDown");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeKeyPress, "KeyPress");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeKeyUp, "KeyUp");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMaster + visEvtAdd, "MasterAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMarker, "MarkerEvent");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMaster + (short)Microsoft.Office.Interop
                .Visio.VisEventCodes.visEvtMod, "MasterChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelMasterDel, "MasterDeleteCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeMouseDown, "MouseDown");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeMouseMove, "MouseMove");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeMouseUp, "MouseUp");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeBefForcedFlush,
                "MustFlushScopeBeginning");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeAfterForcedFlush, "MustFlushScopeEnded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtNonePending, "NoEventsPending");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeWinOnAddonKeyMSG,
                "OnKeystrokeMessageForAddon");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtPage + visEvtAdd, "PageAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtPage + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "PageChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelPageDel, "PageDeleteCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelConvertToGroup,
                "QueryCancelConvertToGroup");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelDocClose,
                "QueryCancelDocumentClose");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelMasterDel,
                "QueryCancelMasterDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelPageDel,
                "QueryCancelPageDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelQuit, "QuerCancelQuit");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelSelDel,
                "QueryCancelSelectionDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelStyleDel,
                "QueryCancelStyleDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelSuspend, "QueryCancelSuspend");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelUngroup, "QueryCancelUngroup");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeQueryCancelWinClose,
                "QueryCancelWindowClose");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelQuit, "QuitCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeDocRunning, "RunModeEntered");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeSelAdded, "SelectionAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeWinSelChange, "SelectionChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelSelDel, "SelectionDeleteCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtShape + visEvtAdd, "ShapeAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtShape + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "ShapeChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeShapeExitTextEdit, "ShapeExitedTextEdit");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeShapeParentChange, "ShapeParentChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeShapeDelete, "ShapesDeleted");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtStyle + visEvtAdd, "StyleAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtStyle + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "StyleChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelStyleDel, "StyleDeleteCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelSuspend, "SuspendCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtText + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "TextChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelUngroup, "UngroupCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeViewChanged, "ViewChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtIdle, "VisioIsIdle");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtApp + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtWinActivate, "WindowActivated");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeCancelWinClose, "WindowCloseCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtWindow + visEvtAdd, "WindowOpened");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtWindow + (short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtMod, "WindowChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio
                .VisEventCodes.visEvtCodeWinPageTurn, "WindowTurnedToPage");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtShapeDataGraphicChanged, "ShapeDataGraphicChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtShapeLinkAdded, "ShapeLinkAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtShapeLinkDeleted, "ShapeLinkDeleted");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtRemoveHiddenInformation, "RemoveHiddenInformation");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeCancelSelGroup, "GroupCanceled");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeQueryCancelSelGroup, "QueryCancelGroup");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtDataRecordset + visEvtAdd, "DataRecordsetAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtDataRecordset + (short)Microsoft.Office.
                Interop.Visio.VisEventCodes.visEvtDel, "BeforeDataRecordsetDelete");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtDataRecordset + (short)Microsoft.Office.
                Interop.Visio.VisEventCodes.visEvtMod, "DataRecordsetChanged");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeCalloutRelationshipAdded, "CalloutRelationshipAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeCalloutRelationshipDeleted, "CalloutRelationshipDeleted");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeContainerRelationshipAdded, "ContainerRelationshipAdded");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeContainerRelationshipDeleted, "ContainerRelationshipDeleted");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeSelectionMovedToSubprocess, "SelectionMovedToSubprocess");
            addEventDescription((short)Microsoft.Office.Interop.Visio.
                VisEventCodes.visEvtCodeRuleSetValidated, "RuleSetValidated");

        }
    }


}
