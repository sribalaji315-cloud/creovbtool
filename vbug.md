| Creo®  | Parametric |       |
| ------ | ---------- | ----- |
| VB API | User’s     | Guide |
9.0.2.0

Copyright©2023PTCInc.and/orItsSubsidiaryCompanies.AllRightsReserved.
CopyrightforPTCsoftwareproductsiswithPTCInc.anditssubsidiarycompanies(collectively“PTC”),and
theirrespectivelicensors.Thissoftwareisprovidedunderwrittenlicenseorotheragreement,contains
valuabletradesecretsandproprietaryinformation,andisprotectedbythecopyrightlawsoftheUnitedStates
andothercountries.Itmaynotbecopiedordistributedinanyformormedium,disclosedtothirdparties,or
usedinanymannernotprovidedforintheapplicableagreementexceptwithwrittenpriorapprovalfrom
PTC.MoreinformationregardingthirdpartycopyrightsandtrademarksandalistofPTC’sregistered
copyrights,trademarks,andpatentscanbeviewedhere:www.ptc.com/support/go/copyright-and-trademarks
UserandtrainingguidesandrelateddocumentationfromPTCarealsosubjecttothecopyrightlawsofthe
UnitedStatesandothercountriesandareprovidedunderalicenseagreementthatrestrictscopying,
disclosure,anduseofsuchdocumentation.PTCherebygrantstothelicensedsoftwareusertherighttomake
copiesofproductdocumentationandguidesinprintedform,butonlyforinternal/personaluseandin
accordancewiththelicenseagreementunderwhichtheapplicablesoftwareislicensed.Anycopymadeshall
includethePTCcopyrightnoticeandanyotherproprietarynoticeprovidedbyPTC.Notethattraining
materialsmaynotbecopiedwithouttheexpresswrittenconsentofPTC.Thisdocumentationmaynotbe
disclosed,transferred,modified,orreducedtoanyform,includingelectronicmedia,ortransmittedormade
publiclyavailablebyanymeanswithoutthepriorwrittenconsentofPTCandnoauthorizationisgrantedto
makecopiesforsuchpurposes.
UNITEDSTATESGOVERNMENTRIGHTS
PTCsoftwareproductsandsoftwaredocumentationare“commercialitems”asthattermisdefinedat48C.F.
R.2.101.PursuanttoFederalAcquisitionRegulation(FAR)12.212(a)-(b)(ComputerSoftware)(MAY2014)
forcivilianagenciesortheDefenseFederalAcquisitionRegulationSupplement(DFARS)at227.7202-1(a)
(Policy)and227.7202-3(a)(Rightsincommercialcomputersoftwareorcommercialcomputersoftware
documentation)(FEB2014)fortheDepartmentofDefense,PTCsoftwareproductsandsoftware
documentationareprovidedtotheU.S.GovernmentunderthePTCcommerciallicenseagreement.Use,
duplicationordisclosurebytheU.S.Governmentissubjectsolelytothetermsandconditionssetforthinthe
applicablePTCsoftwarelicenseagreement.
PTCInc.,121SeaportBlvd,Boston,MA02210USA

Contents
AboutThisGuide........................................................................................................8
OverviewoftheVBAPI..............................................................................................11
Introduction........................................................................................................12
GettingStarted...................................................................................................13
ObjectTypes......................................................................................................15
ProgrammingConsiderations..............................................................................23
VBAPIFundamentals:ControllingCreoParametric.....................................................31
Overview............................................................................................................32
SimpleAsynchronousMode................................................................................32
StartingandStoppingCreoParametric ................................................................32
ConnectingtoaCreoParametricProcess............................................................35
FullAsynchronousMode.....................................................................................37
TroubleshootingVBAPIApplications...................................................................39
TheVBAPIOnlineBrowser.......................................................................................40
OnlineDocumentationVBAPIWizard...................................................................41
SessionObjects........................................................................................................43
OverviewofSessionObjects...............................................................................44
CompatibilityofDeprecatedMethods...................................................................44
Directories.........................................................................................................45
AccessingtheCreoParametricInterface..............................................................48
Selection..................................................................................................................59
InteractiveSelection...........................................................................................60
AccessingSelectionData....................................................................................61
ProgrammaticSelection......................................................................................63
SelectionBuffer..................................................................................................64
RibbonTabs,Groups,andMenuItems.......................................................................67
CreatingRibbonTabs,Groups,andMenuItems...................................................68
AbouttheRibbonDefinitionFile...........................................................................70
LocalizingtheRibbonUserInterfaceCreatedbytheVBAPIApplications...............73
SupportforLegacytheVBAPIApplications..........................................................74
MigrationofLegacyVBAPIApplications..............................................................75
Menus,Commands,andPop-upMenus.....................................................................76
Introduction........................................................................................................77
MenuBarDefinitions...........................................................................................77
MenusButtonsandMenus..................................................................................77
DesignatingCommands......................................................................................81
Pop-upMenus....................................................................................................83
3

Models.....................................................................................................................87
OverviewofModelObjects..................................................................................88
GettingaModelObject........................................................................................88
ModelDescriptors...............................................................................................88
RetrievingModels...............................................................................................89
ModelInformation...............................................................................................91
ModelOperations...............................................................................................94
RunningCreoModelCHECK...............................................................................96
Drawings................................................................................................................105
OverviewofDrawingsintheVBAPI...................................................................106
CreatingDrawingsfromTemplates.....................................................................106
ObtainingDrawingModels................................................................................108
DrawingInformation..........................................................................................108
DrawingOperations..........................................................................................109
DrawingSheets................................................................................................110
DrawingViews.................................................................................................113
DrawingDimensions.........................................................................................118
DrawingTables.................................................................................................124
DetailItems......................................................................................................132
DetailEntities...................................................................................................134
OLEObjects.....................................................................................................136
DetailNotes.....................................................................................................136
DetailGroups...................................................................................................140
DetailSymbols.................................................................................................141
DetailAttachments............................................................................................151
Solid.......................................................................................................................154
GettingaSolidObject.......................................................................................155
SolidInformation..............................................................................................155
SolidOperations...............................................................................................155
SolidUnits........................................................................................................158
MassProperties...............................................................................................164
Annotations......................................................................................................166
CrossSections.................................................................................................166
Materials..........................................................................................................167
SolidBodies............................................................................................................173
SolidBodyInformation......................................................................................174
WindowsandViews................................................................................................176
Windows..........................................................................................................177
EmbeddedBrowser..........................................................................................179
Views...............................................................................................................180
CoordinateSystemsandTransformations..........................................................181
ModelItem..............................................................................................................188
SolidGeometryTraversal..................................................................................189
GettingModelItemObjects................................................................................189
ModelItemInformation......................................................................................190
4 Creo®ParametricVBAPIUser’sGuide

DuplicatingModelItems.....................................................................................191
LayerObjects...................................................................................................191
Features.................................................................................................................193
AccesstoFeatures...........................................................................................194
FeatureInformation..........................................................................................195
FeatureOperations...........................................................................................196
FeatureGroupsandPatterns.............................................................................198
UserDefinedFeatures......................................................................................200
CreatingFeaturesfromUDFs............................................................................201
DatumFeatures......................................................................................................210
DatumPlaneFeatures......................................................................................211
DatumAxisFeatures........................................................................................213
GeneralDatumPointFeatures..........................................................................214
DatumCoordinateSystemFeatures...................................................................215
GeometryEvaluation...............................................................................................219
GeometryTraversal..........................................................................................220
CurvesandEdges............................................................................................221
Contours..........................................................................................................224
Surfaces..........................................................................................................225
Axes,CoordinateSystems,andPoints...............................................................229
Interference......................................................................................................229
DimensionsandParameters....................................................................................232
Overview..........................................................................................................233
TheParamValueObject....................................................................................233
ParameterObjects............................................................................................234
DimensionObjects............................................................................................240
Relations................................................................................................................243
AccessingRelations.........................................................................................244
AccessingPostRegenerationRelations.............................................................244
AddingaCustomizedFunctiontotheRelationsDialogBoxinCreo
Parametric ...................................................................................................245
AssembliesandComponents...................................................................................249
StructureofAssembliesandAssemblyObjects...................................................250
AssemblingComponents..................................................................................255
RedefiningandReroutingAssemblyComponents...............................................259
ExplodedAssemblies........................................................................................259
SkeletonModels...............................................................................................260
FamilyTables..........................................................................................................262
WorkingwithFamilyTables...............................................................................263
CreatingFamilyTableInstances........................................................................265
CreatingFamilyTableColumns.........................................................................265
ActionListeners......................................................................................................267
TheVBAPIActionListeners..............................................................................268
ActionSources.................................................................................................268
Contents 5

TypesofActionListeners..................................................................................269
CancellinganActionListenerOperation..............................................................275
Interface.................................................................................................................276
ExportingFilesand2DModels..........................................................................277
ExportingtoPDFandU3D................................................................................284
Exporting3DGeometry.....................................................................................291
ShrinkwrapExport............................................................................................294
ImportingFiles..................................................................................................298
Importing3DGeometry.....................................................................................301
PlottingFiles....................................................................................................303
PrintingFiles....................................................................................................304
SolidOperations...............................................................................................311
WindowOperations..........................................................................................312
SimplifiedRepresentations......................................................................................314
Overview..........................................................................................................315
RetrievingSimplifiedRepresentations................................................................316
CreatingandDeletingSimplifiedRepresentations...............................................317
ExtractingInformationAboutSimplifiedRepresentations.....................................317
ModifyingSimplifiedRepresentations.................................................................318
SimplifiedRepresentationUtilities......................................................................320
TaskBasedApplicationLibraries..............................................................................322
ManagingApplicationArguments.......................................................................323
LaunchingaCreoTOOLKITDLL.......................................................................324
LaunchingTasksfromJ-LinkTaskLibraries........................................................326
Graphics.................................................................................................................328
Overview..........................................................................................................329
GettingMouseInput..........................................................................................329
DisplayingGraphics..........................................................................................330
DisplayListsandGraphics................................................................................333
ExternalData..........................................................................................................335
ExternalData...................................................................................................336
Exceptions.......................................................................................................339
WindchillConnectivityAPIs......................................................................................341
Introduction......................................................................................................342
AccessingaWindchillServerfromaCreoParametricSession.............................342
AccessingWorkspaces.....................................................................................345
WorkflowtoRegisteraServer............................................................................346
AliasedURL.....................................................................................................347
ServerOperations............................................................................................348
UtilityAPIs.......................................................................................................359
SampleBatchWorkflow....................................................................................359
TechnicalSummaryofChangesforCreo9.0.0.0.......................................................361
NewFunctions.................................................................................................362
MiscellaneousTechnicalChanges.....................................................................362
6 Creo®ParametricVBAPIUser’sGuide

FullVersionofCreo®ParametricVBAPIReleaseNotes.....................................363
SampleApplications................................................................................................364
InstallingtheVBAPI.........................................................................................365
SampleApplications.........................................................................................365
GeometryTraversal.................................................................................................370
Example1........................................................................................................371
Example2........................................................................................................371
Example3........................................................................................................372
Example4........................................................................................................372
Example5........................................................................................................373
GeometryRepresentations......................................................................................374
SurfaceParameterization..................................................................................375
EdgeandCurveParameterization.....................................................................384
Index......................................................................................................................388
Contents 7

1
|     |     | About | This | Guide |
| --- | --- | ----- | ---- | ----- |
Thissectioncontains informationaboutthecontents ofthis user’s guideandthe
conventionsused.
Purpose
Thismanual describeshowtousethe VBAPI, aVisualBasictoolkit forCreo
Parametric.The VBAPI makespossiblethe developmentofVisualBasic
programsthataccesstheinternalcomponentsofaCreo Parametricsession, to
| customizeCreo Parametricmodels. |     |     |     |     |
| ------------------------------- | --- | --- | --- | --- |
Note
The VBAPI issupportedonly with CreoParametric.It isnotsupported with
| theotherCreo | applications. |     |     |     |
| ------------ | ------------- | --- | --- | --- |
Audience
Thismanual isintendedforexperienced CreoParametricuserswhoarefamiliar
| withVisualBasic | oranotherobject-orientedlanguage. |     |     |     |
| --------------- | --------------------------------- | --- | --- | --- |
Prerequisites
| Thismanual assumesyouhavethefollowingknowledge: |     |     |     |     |
| ----------------------------------------------- | --- | --- | --- | --- |
• CreoParametric
| • VisualBasicfor        | Applications(Office  | macros)        |     |     |
| ----------------------- | -------------------- | -------------- | --- | --- |
| • VisualBasic.NET       | 2010or above         |                |     |     |
| • Otherlanguageswiththe | built-incapabilityto | useCOMservers: |     |     |
○ JavaScript
| 8   |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

○ VB.Script
○ C++
○ C#
Documentation
| Thedocumentationfor  | VisualBasicsAPIs | includesthe | following: |
| -------------------- | ---------------- | ----------- | ---------- |
| • VBAPI User’s Guide |                  |             |            |
• Anonlinebrowserthatdescribesthe syntaxofthe VisualBasicfunctions and
providesalinkto theonlineversionof thismanual.The onlineversion ofthe
documentationisupdatedmorefrequently thanthe printedversion.Ifthereare
| anydiscrepancies,the | onlineversion | isthe correctone. |     |
| -------------------- | ------------- | ----------------- | --- |
Conventions
Thefollowingtablelists conventionsandtermsusedthroughout thisbook.
| Convention |     | Description                             |     |
| ---------- | --- | --------------------------------------- | --- |
| UPPERCASE  |     | CreoParametric-typemenuname(forexample, |     |
PART).
| Boldface |     | Windows-typemenunameormenuordialogbox |     |
| -------- | --- | ------------------------------------- | --- |
option(forexample,View),orutility.Boldface
fontisalsousedforkeywords,VBmethods,names
ofdialogboxbuttons,andCreoParametric
commands.
| Monospace(Courier) |     | Codesamplesappearincourierfontlike |     |
| ------------------ | --- | ---------------------------------- | --- |
this.VBaspects(methods,classes,datatypes,
objectnames,andsoon)alsoappearinCourierfont.
| Emphasis |     | Importantinformationappearsinitalicslikethis. |     |
| -------- | --- | --------------------------------------------- | --- |
Italicfontisalsousedforfilenamesanduniform
resourcelocators(URLs).
| Choose |     | Highlightamenuoptionbyplacingthearrowcursor |     |
| ------ | --- | ------------------------------------------- | --- |
ontheoptionandpressingtheleftmousebutton.
| Select |     | Asynonymfor“choose”asabove,Selectalso |     |
| ------ | --- | ------------------------------------- | --- |
describestheactionsofselectingelementsona
modelandcheckingboxes.
| Element |     | Anelementdescribesredefinablecharacteristicsofa |     |
| ------- | --- | ----------------------------------------------- | --- |
featureinamodel.
| Mode |     | AnenvironmentinCreoParametricinwhichyou |     |
| ---- | --- | --------------------------------------- | --- |
canperformagroupofcloselyrelatedfunctions
(Drawing,forexample).
| Model |     | Anassembly,part,drawing,format,notebook,case |     |
| ----- | --- | -------------------------------------------- | --- |
study,sketch,andsoon.
| Option |     | Aniteminamenuoranentryinaconfigurationfile |     |
| ------ | --- | ------------------------------------------ | --- |
orasetupfile.
| Solid |     | Apartoranassembly. |     |
| ----- | --- | ------------------ | --- |
AboutThisGuide 9

| <creo_loadpoint> |     | ThelocationwheretheCreoapplicationsare |     |
| ---------------- | --- | -------------------------------------- | --- |
installed,forexample,C:\ProgramFiles\
PTC\Creo1.0.
| <creo_vbapi_loadpoint> |     | ThelocationwheretheVBAPIapplicationis |     |
| ---------------------- | --- | ------------------------------------- | --- |
installed,thatis,<creo_loadpoint>\
<datecode>\CommonFiles\vbapi.
Note
• Important informationthatshouldnotbe overlookedappears in noteslike
this.
• Allreferencesto mouseclicks assumeuseof aright-handedmouse.
| SoftwareProductConcernsand |     | Documentation | Comments |
| -------------------------- | --- | ------------- | -------- |
Forresourcesandservicesto helpyouwithPTC software products,seethe PTC
CustomerService Guide.It includesinstructionsfor usingthe World Wide Web or
| faxtransmissionsfor | customer support. |     |     |
| ------------------- | ----------------- | --- | --- |
Inregardto documentation,PTCwelcomesyoursuggestionsandcomments.You
| cansendfeedback | inthe followingways: |                             |     |
| --------------- | -------------------- | --------------------------- | --- |
| • Sendcomments  | electronicallyto     | MCAD-documentation@ptc.com. |     |
• Filloutandmailthe PTCDocumentationSurveyin thecustomer service
guide.
| 10  |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | ------------------------------- |

2
|     | Overview |     | of the | VB API |
| --- | -------- | --- | ------ | ------ |
Introduction...............................................................................................................12
GettingStarted..........................................................................................................13
ObjectTypes.............................................................................................................15
ProgrammingConsiderations.....................................................................................23
| Thischapterprovidesan | overviewof | theVB APIs. |     |     |
| --------------------- | ---------- | ----------- | --- | --- |
11

Introduction
TheVB APIfor CreoParametricis an asynchronousapplicationthatcan be used
fromanyCOM-enabledapplicationincludingVisualBasic.NET(VB.NET),
VisualBasicfor Applications(VBA),andexternalInternetExplorer instances
usingscripting.
| Visual          | Basic.NETApplications |                       |     |     |
| --------------- | --------------------- | --------------------- | --- | --- |
| You canusetheVB |                       | APIfor CreoParametric |     | to: |
• CreateaVB.NETform capable ofstartingor connectingto Creo Parametric
non-graphically,accepting userinputsanddrivingmodelmodificationsor
deliverables.
• CreateaVB.NETapplicationthatmayormaynothaveitsownUserInterface
(UI).The applicationshouldbe ableto establish oneormoreCreo Parametric
UIorevent listenersin session, andprocessthoseeventsusing VB.NETcode.
| Visual | Basic | for Applications |     |     |
| ------ | ----- | ---------------- | --- | --- |
TheVB APIsprovidesupport foraccessing CreoParametricfrom VisualBasic-
enabledproductssuchasMicrosoft Excel,Microsoft Word, orMicrosoft Access.
TheCOMinterfaceis providedto controlCreoParametric asynchronouslyand
| usethe | PFCAPI'sto | accessitsproperties. |     |     |
| ------ | ---------- | -------------------- | --- | --- |
You canalsoaccess datafromOLE objectsembeddedin CreoParametric.The
OLEobjectscaninclude VB codethatcan beused todrivethe modelfromwhich
theobjectiscontained.
| Limitations        |     | of the VB    | API                   |     |
| ------------------ | --- | ------------ | --------------------- | --- |
| TheasynchronousCOM |     | serverhasthe | followinglimitations: |     |
• APIcallsto Creo Parametricshouldbe madeonly froma singlethread.Other
threadscan processnonCreo Parametricdataandsetdatatobe seen bythe
CreoParametricthread, butonly onethreadcan communicatewith Creo
Parametric.
• Onlyoneactive connectioncan be madeto asingleCreo Parametricsessionat
onetime.
| 12  |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

| Getting |     | Started |             |     |     |     |     |     |
| ------- | --- | ------- | ----------- | --- | --- | --- | --- | --- |
| Setting | Up  | a VB    | Application |     |     |     |     |     |
Foryourapplicationto communicatewith Creo Parametric,youmustsetthe
PRO_COMM_MSG_EXEenvironmentvariableto thefull pathoftheexecutable,
| pro_comm_msg.exe.Typically,                |     |     |                       | thepathto |     | theexecutableis | <creo_     |         |
| ------------------------------------------ | --- | --- | --------------------- | --------- | --- | --------------- | ---------- | ------- |
| loadpoint>\<datecode>\CommonFiles\<machine |     |     |                       |           |     |                 | type>\obj\ |         |
| pro_comm_msg.exe,wheremachinetypeisi486_nt |     |     |                       |           |     |                 | for32-bit  | Windows |
| andx86e_win64for64-bit                     |     |     | Windowsinstallations. |           |     |                 |            |         |
SetPRO_COMM_MSG_EXEas:
| 1. ClickStart>           |                   | Settings> | ControlPanel.                  |                  |     |     |     |     |
| ------------------------ | ----------------- | --------- | ------------------------------ | ---------------- | --- | --- | --- | --- |
| 2. ClickSystem.The       |                   |           | SystemPropertieswindowsopens.  |                  |     |     |     |     |
| 3. Inthe                 | Advancedtab,click |           | theEnvironmentVariablesbutton. |                  |     |     |     |     |
| 4. AddPRO_COMM_MSG_EXEto |                   |           |                                | Systemvariables. |     |     |     |     |
| Registering              |                   | the       | COM Server                     |                  |     |     |     |     |
Toregisterthe COMserver, runthevb_api_register.bat filelocatedat
<creo_loadpoint>/Parametric/bin.
Tounregisterthe COMserver, runthevb_api_unregister.bat filelocated
at<creo_loadpoint>/Parametric/bin.
AftertheCOM serveris registeredwith thesystem,wheneveran applicationtries
toaccessthe typescontainedin thisservertheserverstartsautomatically.By
default,Windowsstartsservicessuchaspfclscom.exein
the Windowssystem
directory(c:\winnt\system_32). Becausetheserverwillalsostartnew
sessionsofCreo Parametricfromthe processworkingdirectory, youmaywant to
controltheserverrundirectory.You canconfigure theserverto startin aspecific
directoryby settingthesystem environmentvariablePFCLS_START_DIR to any
| existingdirectoryonyour |         |                | computer. |     |       |                |     |      |
| ----------------------- | ------- | -------------- | --------- | --- | ----- | -------------- | --- | ---- |
| Setting                 | Project | References     |           | for | the   | VB API         |     |      |
| Setthereferenceto       |         | CreoParametric |           |     | VBAPI | TypeLibraryfor |     | Creo |
ParametricCreo+through yourproject.Inthe VBAenvironmentsetthis reference
asfollows:
1. ClickTools▶▶References.
| 2. Checkthe       |     | boxforCreoParametricVB |       |                  |     | APITypeLibrary |     | forCreo |
| ----------------- | --- | ---------------------- | ----- | ---------------- | --- | -------------- | --- | ------- |
| Parametricasshown |     |                        | inthe | followingfigure. |     |                |     |         |
OverviewoftheVBAPI 13

| IntheVB.NETenvironment, setthisreferenceasfollows: |     |     |
| -------------------------------------------------- | --- | --- |
1. ClickProject▶▶Properties▶▶AddReference▶▶COM.
| 2. SelectCreoParametricVB | APITypeLibraryfor | Creo |
| ------------------------- | ----------------- | ---- |
| Parametricasshown inthe   | followingfigure.  |      |
14 Creo®ParametricVBAPIUser’sGuide

Object Types
TheVB APIis madeup ofanumberof classesin manymodules.Thefollowing
arethemainclass types:
| • CreoParametric-RelatedClasses—Containunique |     |     | methodsand |
| --------------------------------------------- | --- | --- | ---------- |
propertiesthataredirectlyrelated to thefunctionsin Creo Parametric.Seethe
sectionCreoParametric-Related Classeson page16 for moreinformation.
• CompactDataClasses—Classescontainingdataneededasargumentsto
| someVBmethods.Seethe | section,CompactDataClasses |     | onpage18,for |
| -------------------- | -------------------------- | --- | ------------ |
additionalinformation.
• UnionClasses—Classeswith apotential to containmultiple typesof
| values.SeethesectionUnionson | page18for | additionalinformation. |     |
| ---------------------------- | --------- | ---------------------- | --- |
• SequenceClasses—Expandablearraysofobjectsorprimitivedatatypes.
| SeethesectionSequenceson | page19 | for moreinformation. |     |
| ------------------------ | ------ | -------------------- | --- |
OverviewoftheVBAPI 15

• ArrayClasses—Arraysthatare limitedtoa certainsize.See thesection
| Arrayson | page20for | moreinformation. |     |     |     |
| -------- | --------- | ---------------- | --- | --- | --- |
• Enumeration Classes—Enumeratedtypes,whichlist arestrictedand
validsetofoptions for theproperty. SeethesectionEnumerationClasses on
| page20for                                      | moreinformation. |     |     |     |              |
| ---------------------------------------------- | ---------------- | --- | --- | --- | ------------ |
| • Module-LevelClasses—Containstaticmethodsused |                  |     |     |     | toinitialize |
certainVB objects.See thesectionModule-LevelClasseson page21for more
information.
| • ActionListenerClasses—Enableyouto |     |     |     | specifycodethatwillrun |     |
| ----------------------------------- | --- | --- | --- | ---------------------- | --- |
onlyif certainevents inCreo Parametrictakeplace.Seethe sectionAction
| Listenerson | page21for | moreinformation. |     |     |     |
| ----------- | --------- | ---------------- | --- | --- | --- |
Eachclassshares specificrules regarding initialization,attributes,methods,
inheritance,orexceptions.The followingsectionsdescribetheseclassesin detail.
| Creo Parametric                  |     | -Related | Classes |                            |     |
| -------------------------------- | --- | -------- | ------- | -------------------------- | --- |
| TheCreoParametric-RelatedClasses |     |          |         | containmethodsthatdirectly |     |
manipulateobjectsin Creo Parametric.Examples oftheseobjectsincludemodels,
features,andparameters.
Initialization
You cannotconstruct oneof theseobjectsusingthe keywordNew.Instead, you
shouldobtainthe handletoa CreoParametric-relatedobjectby creatingorlisting
| thatobjectwith                          | amethod             | onthe | parentobjectin | thehierarchy.     |     |
| --------------------------------------- | ------------------- | ----- | -------------- | ----------------- | --- |
| Forexample,IpfcBaseSession.CurrentModel |                     |       |                | returnsaIpfcModel |     |
| objectsetto                             | the currentmodeland |       |                |                   |     |
IpfcParameterOwner.CreateParam()
returnsanewlycreated parameter
| objectfor | manipulation. |     |     |     |     |
| --------- | ------------- | --- | --- | --- | --- |
Properties
Propertieswithin Creo Parametric-relatedobjectsare directlyaccessible. Some
attributesthathavebeendesignated asread-only canbe accessedbut notmodified
| bytheVB | API. |     |     |     |     |
| ------- | ---- | --- | --- | --- | --- |
Methods
You mustinvoke methodsfrom theobjectin questionand firstinitializethat
| object.For         | example,thefollowingcallsareillegal: |     |              |         |     |
| ------------------ | ------------------------------------ | --- | ------------ | ------- | --- |
| Dim window         | as pfcls.IpfcWindow;                 |     |              |         |     |
| window.Activate(); |                                      |     | ‘ The window | has not | yet |
‘ been initialized.
| Repaint(); |     |     | ‘ There | is no invoking                  | object. |
| ---------- | --- | --- | ------- | ------------------------------- | ------- |
| 16         |     |     |         | Creo®ParametricVBAPIUser’sGuide |         |

Thefollowingcallsarelegal:
| Dim | window As       | Pfcls.IpfcWindow               |                           |     |     |
| --- | --------------- | ------------------------------ | ------------------------- | --- | --- |
| Dim | session as      | pfcls.IpfcSession              |                           |     |     |
| Dim | asyncConnection | as                             | pfcls.IpfcAsyncConnection |     |     |
| Dim | Casync as       | New pfcls.CCpfcAsyncConnection |                           |     |     |
asyncConnection = Casync.Connect (DBNull.Value, DBNull.Value,
| DBNull.Value, |                            | DBNull.Value) |     |       |                  |
| ------------- | -------------------------- | ------------- | --- | ----- | ---------------- |
| session       | = asyncConnection.Session; |               |     |       |                  |
| window        | = session.CurrentWindow;   |               |     | ' You | have initialized |
' the window object.
window.Activate()
window.Repaint()
Inheritance
ManyCreo Parametric-relatedobjectsinheritmethodsfrom otherinterfaces.In
VB.NETandVBA,youmusthavean objectofthe correcttypefor thecompiler
| andIDEtoresolve |     | themethods | youwishto | call.For | example,an |
| --------------- | --- | ---------- | --------- | -------- | ---------- |
IpfcComponentFeatobjectcould usethe methodsandpropertiesas follows:
• IpfcObject
• IpfcChild
• IpfcActionSource
• IpfcModelItem
• IpfcFeature
• IpfcComponentFeat
Thefollowingaretheapproachesto using an object’s inheritedmethods:
1. You cancodethemethod calldirectlyeventhough itisnotavailablein
Intellisense.
|     | Dim componentFeat | as      | pfcls.IpfcComponentFeat    |     |     |
| --- | ----------------- | ------- | -------------------------- | --- | --- |
|     | MsgBox ("Feature  | number: | " & componentFeat.Number); |     |     |
Note
|     | This worksin | VB.NETbutis | likelytoresult |     | ina compilationerrorin |
| --- | ------------ | ----------- | -------------- | --- | ---------------------- |
VBA.
2. You cancreateanotherobjectof theappropriatetypeandassignittheobject
|     | handle,andthencallthe |                   | requiredmethod.         |     |     |
| --- | --------------------- | ----------------- | ----------------------- | --- | --- |
|     | Dim componentFeat     | as                | pfcls.IpfcComponentFeat |     |     |
|     | Dim feat as           | pfcls.IpfcFeature |                         |     |     |
|     | feat = componentFeat  |                   |                         |     |     |
OverviewoftheVBAPI 17

| MsgBox  | ("Feature |     | number: | "   | & feat.Number); |     |     |
| ------- | --------- | --- | ------- | --- | --------------- | --- | --- |
| Compact | Data      |     | Classes |     |                 |     |     |
Compactdataclasses aredata-onlyclasses.Theyareusedfor argumentsand
returnvaluesfor someVBAPI methods.Theydonot representactualobjectsin
CreoParametric.Otherthanadifferencein howtheyareinitialized,compactdata
| classeshavesimilarrequirements |     |     |     | toCreo | Parametric-relatedclasses. |     |     |
| ------------------------------ | --- | --- | --- | ------ | -------------------------- | --- | --- |
Initialization
You cancreatethesecompactdataobjectsusing adesignatedCreatemethod
whichresidesonthe CCversionof thecompactclass.Youinstantiate theCCclass
New.
| objectwith | thekeyword |     |     |     |     |     |     |
| ---------- | ---------- | --- | --- | --- | --- | --- | --- |
Forexample,
| 'Class object,      |     | owns                        | Create() |                      |     |     |     |
| ------------------- | --- | --------------------------- | -------- | -------------------- | --- | --- | --- |
| Dim tableCellCreate |     |                             | As New   | pfcls.CCpfcTableCell |     |     |     |
| Dim tableCell       |     | As pfcls.IpfcTableCell      |          |                      |     |     |     |
| Set tableCell       |     | = tableCellCreate.Create(1, |          |                      |     | 1)  |     |
Unions
Unionsare classescontainingpotentiallyseveraldifferent valuetypes. Every
unionhasa discriminatorpropertywith thepredefinedname,discr.This
propertyreturnsavalue identifyingthetypeofdatathattheunionobjectholds.
Foreachunionmember, aseparate propertyisusedto accessthe differentdata
types.It isillegaltoattempt toread anypropertyexcepttheone thatmatchesthe
valuereturnedfrom thediscriminator.However, anypropertythatswitchesthe
| discriminatortothe |                | newvaluetypecanbe |     |       | modified. |     |     |
| ------------------ | -------------- | ----------------- | --- | ----- | --------- | --- | --- |
| Thefollowingisan   |                | exampleofa        |     | VBAPI | union:    |     |     |
| Interface          | IpfcParamValue |                   |     |       |           |     |     |
--------------------------------------------------------
Description
| This class |     | describes | the | value | of  | the parameter. |     |
| ---------- | --- | --------- | --- | ----- | --- | -------------- | --- |
--------------------------------------------------------
| Union    | Discriminant |       |                    |     |        |            |     |
| -------- | ------------ | ----- | ------------------ | --- | ------ | ---------- | --- |
| Property | discr        | as    | IpfcParamValueType |     |        | [readonly] |     |
| Returns  | the          | union | discriminant       |     | value. |            |     |
--------------------------------------------------------
| Property | Summary   |     |      |                   |     |         |           |
| -------- | --------- | --- | ---- | ----------------- | --- | ------- | --------- |
| Property | BoolValue |     | as   | Boolean           |     |         |           |
| If the   | parameter |     | type | is PARAM_BOOLEAN, |     | this is | a Boolean |
value.
| Property | DoubleValue |     | as  | Double |     |     |     |
| -------- | ----------- | --- | --- | ------ | --- | --- | --- |
If the parameter type is PARAM_DOUBLE, this is a double value.
| Property | IntValue |     | as Long |     |     |                                 |     |
| -------- | -------- | --- | ------- | --- | --- | ------------------------------- | --- |
| 18       |          |     |         |     |     | Creo®ParametricVBAPIUser’sGuide |     |

| If the | parameter | type | is  | PARAM_INTEGER, |     | this is | an integer |
| ------ | --------- | ---- | --- | -------------- | --- | ------- | ---------- |
value.
| Property | NoteId | as Long |     |     |     |     |     |
| -------- | ------ | ------- | --- | --- | --- | --- | --- |
If the parameter type is PARAM_NOTE, this is a note identifier.
| Property | StringValue |     | as String |     |     |     |     |
| -------- | ----------- | --- | --------- | --- | --- | --- | --- |
If the parameter type is PARAM_STRING, this is a string value.
Sequences
Sequencesareexpandablearraysof primitivedatatypesorobjects inthe VBAPI.
Allsequenceclasseshavethe samemethodsforadding to andaccessing thearray.
Sequenceclassesare typicallyidentifiedby aplural name,orthesuffixseq.
Initialization
You cancreateinstances oftheseclassesdirectlyby instantiatingtheappropriate
classobject:
Properties
Theread-only Countattribute identifieshowmanymembers arecurrentlyin the
sequence.You mayalsoaccessmembers ofthe sequenceusing theItemproperty
ordirectly:
| Dim model | as IpfcModel |     |     |     |     |     |     |
| --------- | ------------ | --- | --- | --- | --- | --- | --- |
| model =   | models (0)   |     |     |     |     |     |     |
Methods
Sequenceobjectsalways containthesamemethods.Use thefollowingmethodsto
| accessthecontents          | ofthesequence: |     |           |                    |          |       |     |
| -------------------------- | -------------- | --- | --------- | ------------------ | -------- | ----- | --- |
| • Append()—Addsanewitem    |                |     |           | to theendof        | thearray |       |     |
| • Clear()—Removesall       |                |     | itemsfrom |                    | thearray |       |     |
| • Insert()—Insertsanewitem |                |     |           | atanylocationofthe |          | array |     |
• InsertSeq()—Inserts thecontentsof asequence ofitems atany location
ofthearray
| • Set()—Assignsoneitem |     |     | in            | thearraytothe |         | inputitem |     |
| ---------------------- | --- | --- | ------------- | ------------- | ------- | --------- | --- |
| • Remove()—Removes     |     |     | arangeofitems |               | fromthe | array     |     |
OverviewoftheVBAPI 19

Inheritance
Sequenceclassesdo notinheritfrom anyotherVB APIclasses.Therefore, you
cannotusesequence objectsasargumentswhereanyothertypeofVB APIobject
isexpected,includingothertypesofsequences.For example,ifyouhavealist of
IpfcModelItemsthathappento be features,youcannotusethe sequenceas if
| itwere asequence          | ofIpfcFeatures.                            |     |                      |     |       |
| ------------------------- | ------------------------------------------ | --- | -------------------- | --- | ----- |
| Toconstruct               | thearrayoffeatures,youmustinserteachmember |     |                      |     | ofthe |
| IpfcModelItemslistintothe |                                            |     | newIpfcFeatureslist. |     |       |
Arrays
Arraysaregroups ofprimitivetypesorobjectsof aspecifiedsize.An arraycan be
one-ortwo- dimensional.Theonlinereferencedocumentationindicatestheexact
sizeofeacharrayclass.
Initialization
You cancreateinstances oftheseclassesdirectlyby instantiatingtheappropriate
classobject:
Properties
You mayread members ofthesequence using theItempropertyordirectlyasan
array:
| Dim point    | as IpfcPoint3D  |            |                |         |     |
| ------------ | --------------- | ---------- | -------------- | ------- | --- |
| Dim matrix   | as IpfcMatrix3D |            |                |         |     |
| MsgBox (“Y   | value           | of point:  | “ & point.Item | (1))    |     |
| MsgBox (“(2, | 2) value        | of matrix: | “ & matrix     | (2, 2)) |     |
Methods
Arrayobjectscontainonly theSet method,whichassignsoneitemin thearrayto
theinputitem.
| Enumeration |     | Classes |     |     |     |
| ----------- | --- | ------- | --- | --- | --- |
IntheVB API,an enumerationclassdefines alimitednumberof valuesthat
correspondto themembers ofthe enumeration.Eachvalue representsan
appropriatetypeandmaybe accessedby name.Inthe EpfcFeatureType
enumerationclass, thevalueEpfcFEATTYPE_HOLE represents aHole featurein
CreoParametric.Enumerationclassesin theVB APIgenerally havenames ofthe
| formEpfcXYZType |     | orEpfcXYZStatus. |     |                                 |     |
| --------------- | --- | ---------------- | --- | ------------------------------- | --- |
| 20              |     |                  |     | Creo®ParametricVBAPIUser’sGuide |     |

Initialization
You candirectlyreferto instance ofthisclass:
Dim type as EpfcFeatureType
type = EpfcFeatureType.EpfcFEATTYPE_HOLE
Attributes
Anenumerationclassis madeup ofconstantinteger properties.The namesof
thesepropertiesare alluppercase anddescribewhat theattribute represents.For
example:
• EpfcPARAM_INTEGER—Avaluein theEpfcParamValueType
enumerationclass thatisused toindicatethata parameterstores an integer
value.
• EpfcITEM_FEATURE—An valuein theEpfcModelItemType
enumerationclass thatisused toindicatethata modelitemisa feature.
Anenumerationclassalways hasan integervalue named<type>_nil, whichis
onemorethanthe highestacceptablenumericalvaluefor thatenumerationclass.
Module-Level Classes
Somemodulesin theVB APIhaveoneclass thatcontainsspecialfunctions used
tocreateandaccess someof theotherclassesin the package.Thesemodule
classeshavethenamingconvention, CM+the nameofthe module,forexample
CMpfcSelect.
Initialization
You cancreateinstances oftheseclassesdirectlyby instantiatingtheappropriate
classobject:
Dim mSelect as New CMpfcSelect
Methods
Module-levelclasses containonly staticmethods usedforinitializing certainVB
APIobjects.
Action Listeners
ActionListeners notifyyouof eventsin Creo Parametric.Theyarealso thebasis
forcustomizationofthe CreoParametricUserInterface. ActionListenersarenot
supportedfrom VBA.
OverviewoftheVBAPI 21

Initialization
InVB.NET, youcancreateandassignan ActionListenerclass asfollows.
Createa classimplementingthe listenerin question.It shoulddefineall the
inheritedmethods,evenifyouwant toonly execute codefor afewof thelistener
methods.Thoseothermethodsshouldbe implementedwith an emptybody.
Theclass shouldalsoimplementtheinterfaceIpfcActionListener, which
hasno methods.
Theclass shouldalsoimplementICIPClientObject.This methoddefines the
objecttypeto theCIPcodein theserver. Thismethod returnsaStringwhichis the
| nameofthe | listenertypeinterface,forexample, |     |     |     |     |     |     |     |     |
| --------- | --------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- |
IpfcSessionActionListener.
| Private | Class                  | ModelEventListener |                              |                                         |     |     |     |           |     |
| ------- | ---------------------- | ------------------ | ---------------------------- | --------------------------------------- | --- | --- | --- | --------- | --- |
|         | Implements             |                    | IpfcModelEventActionListener |                                         |     |     |     |           |     |
|         | Implements             |                    | ICIPClientObject             |                                         |     |     |     |           |     |
|         | Implements             |                    | IpfcActionListener           |                                         |     |     |     |           |     |
|         | Public                 | Function           |                              | GetClientInterfaceName()                |     |     |     | As String | _   |
|         | Implements             |                    |                              | ICIPClientObject.GetClientInterfaceName |     |     |     |           |     |
|         | GetClientInterfaceName |                    |                              |                                         | =   |     |     |           |     |
"IpfcModelEventActionListener"
End Function
'================================================================-
======
| 'Function | :   | OnAfterModelCopy |        |          |             |       |     |              |     |
| --------- | --- | ---------------- | ------ | -------- | ----------- | ----- | --- | ------------ | --- |
| 'Purpose  | :   | This             | method |          | is executed | after |     | successfully |     |
| '         |     | copying          |        | a model. |             |       |     |              |     |
'================================================================-
======
|     | Public                     | Sub | OnAfterModelCopy(ByVal |     |     |            | _FromMdl | As  |     |
| --- | -------------------------- | --- | ---------------------- | --- | --- | ---------- | -------- | --- | --- |
|     | pfcls.IpfcModelDescriptor, |     |                        |     |     | ByVal      | _ToMdl   | As  |     |
|     | pfcls.IpfcModelDescriptor) |     |                        |     |     | Implements |          |     |     |
pfcls.IpfcModelEventActionListener.OnAfterModelCopy
|     | 'Method | Body |     |     |     |     |     |     |     |
| --- | ------- | ---- | --- | --- | --- | --- | --- | --- | --- |
End Sub
'================================================================-
======
| 'Function | :   | OnAfterModelRename |        |          |             |       |     |              |     |
| --------- | --- | ------------------ | ------ | -------- | ----------- | ----- | --- | ------------ | --- |
| 'Purpose  | :   | This               | method |          | is executed | after |     | successfully |     |
| '         |     | renaming           |        | a model. |             |       |     |              |     |
'================================================================-
======
|     | Public                     | Sub | OnAfterModelRename(ByVal |     |     |            | _FromMdl | As  |     |
| --- | -------------------------- | --- | ------------------------ | --- | --- | ---------- | -------- | --- | --- |
|     | pfcls.IpfcModelDescriptor, |     |                          |     |     | ByVal      | _ToMdl   | As  |     |
|     | pfcls.IpfcModelDescriptor) |     |                          |     |     | Implements |          |     |     |
pfcls.IpfcModelEventActionListener.OnAfterModelRename
|     | 'Method | Body |     |     |     |     |                                 |     |     |
| --- | ------- | ---- | --- | --- | --- | --- | ------------------------------- | --- | --- |
| 22  |         |      |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |

End Sub
|     | Public Sub                 | OnAfterModelCopyAll(ByVal |            | _FromMdl | As  |     |
| --- | -------------------------- | ------------------------- | ---------- | -------- | --- | --- |
|     | pfcls.IpfcModelDescriptor, |                           | ByVal      | _ToMdl   | As  |     |
|     | pfcls.IpfcModelDescriptor) |                           | Implements |          |     |     |
pfcls.IpfcModelEventActionListener.OnAfterModelCopyAll
End Sub
|     | Public Sub                 | OnAfterModelDelete(ByVal |            | _Descr | As  |     |
| --- | -------------------------- | ------------------------ | ---------- | ------ | --- | --- |
|     | pfcls.IpfcModelDescriptor) |                          | Implements |        |     |     |
pfcls.IpfcModelEventActionListener.OnAfterModelDelete
End Sub
|     | Public Sub                 | OnAfterModelErase(ByVal |            | _Descr | As  |     |
| --- | -------------------------- | ----------------------- | ---------- | ------ | --- | --- |
|     | pfcls.IpfcModelDescriptor) |                         | Implements |        |     |     |
pfcls.IpfcModelEventActionListener.OnAfterModelErase
End Sub
| End | Class |     |     |     |     |     |
| --- | ----- | --- | --- | --- | --- | --- |
Exceptions
Actionlisteners causemethodsto becalled outsideof yourapplicationstartand
stopmethods.Therefore,youmustinclude exception-handlingcodeinside the
IpfcActionListenerimplementationif youwantto respond toexceptions.
Insomemethodscalledbeforean event,propagatinga
| IpfcXCancelProEAction |     | exceptionoutof | yourmethodwillcancelthe |     |     |     |
| --------------------- | --- | -------------- | ----------------------- | --- | --- | --- |
impendingevent.
| Programming |     | Considerations |     |     |     |     |
| ----------- | --- | -------------- | --- | --- | --- | --- |
Theitems inthis sectionintroduceprogrammingtips andtechniques usedfor
| programmingwith | theVB     | API. |     |     |     |     |
| --------------- | --------- | ---- | --- | --- | --- | --- |
| Application     | Hierarchy |      |     |     |     |     |
Therules ofobjectorientationrequireacertain hierarchyofobjectcreationwhen
youstarta VBapplication. Theapplicationmustiteratedownto the levelofthe
objectyou wantto access.Forexample,to listall thedatumaxes containedin the
| holefeaturesin | all modelsin | session, dothe | following: |     |     |     |
| -------------- | ------------ | -------------- | ---------- | --- | --- | --- |
methodCCpfcAsyncConnection.Connect()
| 1. Usethe          |                      |                        |                      |     | to connectto | an  |
| ------------------ | -------------------- | ---------------------- | -------------------- | --- | ------------ | --- |
| existingsessionof  |                      | CreoParametric.        |                      |     |              |     |
| Dim                | connection           | As IpfcAsyncConnection |                      |     |              |     |
| Dim                | classAsyncConnection | As New                 | CCpfcAsyncConnection |     |              |     |
| OverviewoftheVBAPI |                      |                        |                      |     |              | 23  |

| connection    | = classAsyncConnection.Connect |               |     | (DBNull.Value, |     |
| ------------- | ------------------------------ | ------------- | --- | -------------- | --- |
| DBNull.Value, | DBNull.Value,                  | DBNull.Value) |     |                |     |
2. Getahandleto thesessionof CreoParametricforthecurrentactive
connection:
| Dim session                    | As IpfcBaseSession     |                                      |       |     |     |
| ------------------------------ | ---------------------- | ------------------------------------ | ----- | --- | --- |
| ession                         | = connection.Session   |                                      |       |     |     |
| 3. Getthemodelsthatareloadedin |                        | thesession:                          |       |     |     |
| Dim models                     | As IpfcModels          |                                      |       |     |     |
| models                         | = session.ListModels() |                                      |       |     |     |
| 4. Getthehandleto              | thefirst               | modelin the                          | list: |     |     |
| Dim model                      | As IpfcModel           |                                      |       |     |     |
| model                          | = models[0]            |                                      |       |     |     |
| 5. Getthefeaturemodelitems     |                        | ineachmodel:                         |       |     |     |
| Dim items                      | As IpfcModelItems      |                                      |       |     |     |
| items                          | = model.ListItems      | (EpfcModelItemType.EpfcITEM_FEATURE) |       |     |     |
| 6. Filterout                   | thefeaturesoftypehole: |                                      |       |     |     |
if (feature.FeatType = EpfcFeatureType.EpfcFEATTYPE_HOLE) then
| 7. Getthesubitems | in eachfeaturethatare |     | axes: |     |     |
| ----------------- | --------------------- | --- | ----- | --- | --- |
| Dim axes          | As IpfcModelItems     |     |       |     |     |
axes = feature.ListSubItems (EpfcModelItemType.EpfcITEM_AXIS)
| Optional | Arguments | and Tags |     |     |     |
| -------- | --------- | -------- | --- | --- | --- |
Manymethods inthe VBAPI areshownin theonlinedocumentationashaving
optionalarguments.
| Forexample,theIpfcModelItemOwner.ListItems() |           |     |     | method | takesan |
| -------------------------------------------- | --------- | --- | --- | ------ | ------- |
| optionalType                                 | argument. |     |     |        |         |
IpfcModelItems ListItems (Type as IpfcModelItemType [optional]);
InVB.Net,youcan passthekeyword Nothingin placeof anysuchoptional
useNullin
| argument.InVBA, |     | placeofany | suchoptionalargument.The |     | VB  |
| --------------- | --- | ---------- | ------------------------ | --- | --- |
APImethodsthattakeoptionalargumentsprovidedefaulthandling forNothing
| parameterswhichis | describedin | theonlinedocumentation. |     |     |     |
| ----------------- | ----------- | ----------------------- | --- | --- | --- |
Note
passNothingin
| Youcanonly      |             | place | ofargumentsthatareshown |                                 | inthe |
| --------------- | ----------- | ----- | ----------------------- | ------------------------------- | ----- |
| documentationto | beoptional. |       |                         |                                 |       |
| 24              |             |       |                         | Creo®ParametricVBAPIUser’sGuide |       |

| Optional | Returns |     | for | the VB | API | Methods |     |     |
| -------- | ------- | --- | --- | ------ | --- | ------- | --- | --- |
Somemethodsin theVB APIhavean optionalreturn.Usually thesecorrespond to
lookupmethodsthatmayormaynotfind an objectto return.For example,the
| IpfcBaseSession.GetModel()method |            |       |     |            |      | returnsan | optionalmodel: |     |
| -------------------------------- | ---------- | ----- | --- | ---------- | ---- | --------- | -------------- | --- |
| Function                         | GetModel   | (Name |     | as String, | Type | as        | IpfcModelType) | as  |
| IpfcModel                        | [optional] |       |     |            |      |           |                |     |
TheVB APImightreturnNothing in certaincaseswherethesemethods are
called.Youmustuseappropriatevaluechecks in yourapplicationcodeto handle
thesesituations.
| Parent-Child |     | Relationships |     |     | between |     | the VB | API |
| ------------ | --- | ------------- | --- | --- | ------- | --- | ------ | --- |
Objects
| SomeVBAPI | objectsinheritfrom |     |     | eitherthe | interfaceIpfcParentor |     |     |     |
| --------- | ------------------ | --- | --- | --------- | --------------------- | --- | --- | --- |
IpfcChild.Theseinterfacesareused tomaintain arelationship betweenthe
twoobjects.This hasnothingto dowith object-orientedinheritance,butrather,
refersto therelationship betweentheitems inCreo Parametric.Inthe VBAPI,the
| Childisownedby |     | theParent. |     |     |     |     |     |     |
| -------------- | --- | ---------- | --- | --- | --- | --- | --- | --- |
PropertyIntroduced:
• IpfcChild.DBParent
TheIpfcChild.DBParent propertyreturnstheowner ofthe childobject.The
applicationdevelopermustknowthe expectedtypeof theparent in orderto useit
inlater calls.The followingtablelists parent/childrelationshipsin theVB API.
| Parent      |      |     |                |     | Child                  |        |        |     |
| ----------- | ---- | --- | -------------- | --- | ---------------------- | ------ | ------ | --- |
| IpfcSession |      |     |                |     | IpfcModel              |        |        |     |
| IpfcSession |      |     |                |     | IpfcWindow             |        |        |     |
| IpfcModel   |      |     |                |     | IpfcModelItem          |        |        |     |
| IpfcSolid   |      |     |                |     | IpfcFeature            |        |        |     |
| IpfcModel   |      |     |                |     | IpfcParameter          |        |        |     |
| IpfcModel   |      |     |                |     | IpfcExternalDataAccess |        |        |     |
| IpfcPart    |      |     |                |     | IpfcMaterial           |        |        |     |
| IpfcModel   |      |     |                |     | IpfcView               |        |        |     |
| IpfcModel2D |      |     |                |     | IpfcView2D             |        |        |     |
| IpfcSolid   |      |     |                |     | IpfcXSection           |        |        |     |
| IpfcSession |      |     |                |     | IpfcDll(CreoTOOLKIT)   |        |        |     |
| IpfcSession |      |     |                |     | IpfcJLinkApplication() |        |        |     |
| Run-Time    | Type |     | Identification |     |                        | in the | VB API |     |
TheVB APIandVisualBasicprovideseveralmethodsto identifythe typeof an
object.
OverviewoftheVBAPI 25

Manyclassesprovideread accessto atypeenumeratedclass. Forexample,the
IpfcFeatureclass hasaIpfcFeature.FeatType property,returning a
pfcFeatureTypeenumerationvaluerepresentingthetypeofthe feature.
Baseduponthetype, ausercan recognizethattheIpfcFeature objectis
actuallyaparticular subtype,suchas IpfcComponentFeat,whichis an
assemblycomponent.
| Support | for Embedded |     | OLE | Objects |     |
| ------- | ------------ | --- | --- | ------- | --- |
OLEobjects,whenactivatedby theuser,can includeVB codethatcan beused to
drivethemodelfromwhichthe objectiscontained.Youcan usethe method
IpfcAsyncConnection.GetConnectionId()to getthe connectionID
oftheCreo ParametricsessionthatlaunchedtheprocesscontainingtheOLE
object.Referto thesectionConnectingto aCreo ParametricProcesson page35
| forinformationabout | gettingtheconnectionID. |     |     |     |     |
| ------------------- | ----------------------- | --- | --- | --- | --- |
Theuserapplicationcodeembeddedin theOLEobjectpasses theconnectionID
| stringtoCCpfcConnectionId.Create |     |     |     |     | and |
| -------------------------------- | --- | --- | --- | --- | --- |
IpfcAsyncConnection.ConnectById()to establish theconnection. The
codemaythenobtaintheowner modelof theOLE objectbyretrieving thecurrent
| modelfrom | thesessionusing | standardPFCAPIs. |     |     |     |
| --------- | --------------- | ---------------- | --- | --- | --- |
Forexample,
| Dim ls      | As New pfcls.CCpfcAsyncConnection |     |               |     |               |
| ----------- | --------------------------------- | --- | ------------- | --- | ------------- |
| Dim aC      | As pfcls.IpfcAsyncConnection      |     |               |     |               |
| Dim cId     | As New pfcls.CCpfcConnectionId    |     |               |     |               |
| Dim id      | As pfcls.IpfcConnectionId         |     |               |     |               |
| Dim session | As pfcls.IpfcBaseSession          |     |               |     |               |
| Dim model   | As pfcls.IpfcModel                |     |               |     |               |
| Set id      | = cId.Create(connectionId)        |     |               |     |               |
| Set aC      | = ls.ConnectById(id,              |     | DBNull.Value, |     | DBNull.Value) |
| Set session | = aC.Session                      |     |               |     |               |
| Set model   | = session.CurrentModel            |     |               |     |               |
Exceptions
| AllPFCmethodsthatfailmaythrow |     |     | exceptions |     | as  |
| ----------------------------- | --- | --- | ---------- | --- | --- |
System.Runtime.InteropServices.COMException.
Thetypeofthe exceptioncanbe obtainedfrom theMessageproperty ofthis
exception.
Try
| session.SetConfigOption("no_way", |                 |       |      |                      | "no_how") |
| --------------------------------- | --------------- | ----- | ---- | -------------------- | --------- |
| Catch                             | ex As Exception |       |      |                      |           |
| MsgBox(ex.Message)                |                 | 'Will | show | IpfcXToolkitNotFound |           |
End Try
| 26  |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | --- | ------------------------------- |

Thedescriptionpropertyreturnsthe fullexception descriptionas [Exception
| type];[additional |                                       | details]. | The exceptiontypeisthemodule | and |
| ----------------- | ------------------------------------- | --------- | ---------------------------- | --- |
| exceptionname,for | example,IpfcXToolkitCheckoutConflict. |           |                              |     |
Theadditionaldetailsinclude informationthatwas containedin theexception
whenitwas thrownbythe PFClayer, suchasconflictdescriptionsfor exceptions
causedby serveroperations anderrordetailsfor exceptionsgenerated during
drawingcreation.
PFC Exceptions
ThePFCexceptionsare thrownbythe classesthatmakeup theVB API’s public
| interface.The | followingtabledescribes |     | theseexceptions. |     |
| ------------- | ----------------------- | --- | ---------------- | --- |
| Exception     |                         |     | Purpose          |     |
IpfcXBadExternalData
Anattempttoreadcontentsofanexternaldataobject
thathasbeenterminated.
IpfcXBadGetArgValue Indicatesattempttoreadthewrongtypeofdatafrom
theIpfcArgValueunion.
IpfcXBadGetExternalData Indicatesattempttoreadthewrongtypeofdatafrom
theIpfcExternalDataunion.
IpfcXBadGetParamValue Indicatesattempttoreadthewrongtypeofdatafrom
theIpfcParamValueunion.
IpfcXBadOutlineExcludeType Indicatesaninvalidtypeofitemwaspassedtothe
outlinecalculationmethod.
IpfcXCancelProEAction
ThisexceptiontypewillnotbethrownbyVBAPI
methods,butyoumayinstantiateandthrowthisfrom
certainActionListenermethodstocancelthe
correspondingactioninCreoParametric.
| IpfcXCannotAccess |     |     | ThecontentsofaVBAPIobjectcannotbeaccessed |     |
| ----------------- | --- | --- | ----------------------------------------- | --- |
inthissituation.
| IpfcXEmptyString |     |     | Anemptystringwaspassedtoamethodthatdoes |     |
| ---------------- | --- | --- | --------------------------------------- | --- |
notacceptthistypeofinput.
| IpfcXInvalidEnumValue |     |     | Indicatesaninvalidvalueforaspecified |     |
| --------------------- | --- | --- | ------------------------------------ | --- |
enumerationclass.
| IpfcXInvalidFileName |     |     | Indicatesafilenamepassedtoamethodwas |     |
| -------------------- | --- | --- | ------------------------------------ | --- |
incorrectlystructured.
IpfcXInvalidFileType Indicatesamodeldescriptorcontainedaninvalidfile
typeforarequestedoperation.
IpfcXInvalidModelItem Indicatesthattheitemrequestedtobeusedisno
longerusable(forexample,itmayhavebeen
deleted).
| IpfcXInvalidSelection |     |     | IndicatesthattheIpfcSelectionpassedis |     |
| --------------------- | --- | --- | ------------------------------------- | --- |
invalidorismissinganeededpieceofinformation.
Forexample,itscomponentpath,drawingview,or
parameters.
IpfcXJLinkApplicationException Containsthedetailswhenanattempttocallcodein
anexternalapplicationfailedduetoanexception.
| IpfcXJLinkApplicationInactive |     |     | Unabletooperateontherequested |     |
| ----------------------------- | --- | --- | ----------------------------- | --- |
IpfcJLinkApplicationobjectbecauseithas
beenshutdown.
| OverviewoftheVBAPI |     |     |     | 27  |
| ------------------ | --- | --- | --- | --- |

Exception Purpose
IpfcXJLinkTaskNotFound Indicatesthatthetaskwiththegivennamecouldnot
befoundandrun.
IpfcXModelNotInSession Indicatesthatthemodelisnolongerinsession;it
mayhavebeenerasedordeleted.
IpfcXNegativeNumber Numericargumentwasnegative.
IpfcXNumberTooLarge Numericargumentwastoolarge.
IpfcXProEWasNotConnected TheCreoParametricsessionisnotavailablesothe
operationfailed.
IpfcXSequenceTooLong Sequenceargumentwastoolong.
IpfcXStringTooLong Stringargumentwastoolong.
IpfcXUnimplemented Indicatesunimplementedmethod.
IpfcXUnknownModelExtension
Indicatesthatafileextensiondoesnotmatcha
knownCreoParametricmodeltype.
| Creo TOOLKIT Errors |     |     |
| ------------------- | --- | --- |
TheIpfcXToolkitError exceptiontypesprovideaccessto error codesfrom
CreoTOOLKITfunctions thattheVB APIusesinternallyandto thenames ofthe
functionsreturning sucherrors.IpfcXToolkitError is theexceptionyouare
mostlikelyto encounterbecausetheVB API isbuilton topof CreoTOOLKIT.
Thefollowingtablelists theIpfcXToolkitError typesmethodand shows
thecorrespondingCreo TOOLKIT constantthatindicatesthe causeoftheerror.
| IpfcXToolkitErrorChild | CreoTOOLKITError | #   |
| ---------------------- | ---------------- | --- |
Class
| IpfcXToolkitGeneralEr | PRO_TK_GENERAL_ERROR | -1  |
| --------------------- | -------------------- | --- |
ror
| IpfcXToolkitBadInputs | PRO_TK_BAD_INPUTS | -2  |
| --------------------- | ----------------- | --- |
| IpfcXToolkitUserAbort | PRO_TK_USER_ABORT |     |
-3
| IpfcXToolkitNotFound    | PRO_TK_E_NOT_FOUND     | -4  |
| ----------------------- | ---------------------- | --- |
| IpfcXToolkitFound       | PRO_TK_E_FOUND         | -5  |
| IpfcXToolkitLineTooLong | PRO_TK_LINE_TOO_LONG   | -6  |
| IpfcXToolkitContinue    | PRO_TK_CONTINUE        | -7  |
| IpfcXToolkitBadContext  | PRO_TK_BAD_CONTEXT     | -8  |
| IpfcXToolkitNotImple    | PRO_TK_NOT_IMPLEMENTED | -9  |
mented
| IpfcXToolkitOutOfMemory | PRO_TK_OUT_OF_MEMORY | -10 |
| ----------------------- | -------------------- | --- |
| IpfcXToolkitCommError   | PRO_TK_COMM_ERROR    | -11 |
| IpfcXToolkitNoChange    | PRO_TK_NO_CHANGE     | -12 |
PRO_TK_SUPP_PARENTS
| IpfcXToolkitSuppressed |     | -13 |
| ---------------------- | --- | --- |
Parents
| IpfcXToolkitPickAbove   | PRO_TK_PICK_ABOVE   | -14 |
| ----------------------- | ------------------- | --- |
| IpfcXToolkitInvalidDir  | PRO_TK_INVALID_DIR  | -15 |
| IpfcXToolkitInvalidFile | PRO_TK_INVALID_FILE |     |
-16
| IpfcXToolkitCantWrite   | PRO_TK_CANT_WRITE   | -17                             |
| ----------------------- | ------------------- | ------------------------------- |
| IpfcXToolkitInvalidType | PRO_TK_INVALID_TYPE | -18                             |
| 28                      |                     | Creo®ParametricVBAPIUser’sGuide |

| IpfcXToolkitErrorChild | CreoTOOLKITError | #   |
| ---------------------- | ---------------- | --- |
Class
| IpfcXToolkitInvalidPtr  | PRO_TK_INVALID_PTR | -19 |
| ----------------------- | ------------------ | --- |
| IpfcXToolkitUnavailable | PRO_TK_UNAV_SEC    | -20 |
Section
| IpfcXToolkitInvalidMa | PRO_TK_INVALID_MATRIX | -21 |
| --------------------- | --------------------- | --- |
trix
| IpfcXToolkitInvalidName | PRO_TK_INVALID_NAME | -22 |
| ----------------------- | ------------------- | --- |
| IpfcXToolkitNotExist    | PRO_TK_NOT_EXIST    | -23 |
| IpfcXToolkitCantOpen    | PRO_TK_CANT_OPEN    | -24 |
| IpfcXToolkitAbort       | PRO_TK_ABORT        | -25 |
| IpfcXToolkitNotValid    | PRO_TK_NOT_VALID    | -26 |
| IpfcXToolkitInvalidItem | PRO_TK_INVALID_ITEM | -27 |
PRO_TK_MSG_NOT_FOUND
| IpfcXToolkitMsgNotFound |                      | -28 |
| ----------------------- | -------------------- | --- |
| IpfcXToolkitMsgNoTrans  | PRO_TK_MSG_NO_TRANS  | -29 |
| XToolkitMsgFmtError     | PRO_TK_MSG_FMT_ERROR | -30 |
| IpfcXToolkitMsgUserQuit | PRO_TK_MSG_USER_QUIT | -31 |
| IpfcXToolkitMsgTooLong  | PRO_TK_MSG_TOO_LONG  |     |
-32
| IpfcXToolkitCantAccess | PRO_TK_CANT_ACCESS   | -33 |
| ---------------------- | -------------------- | --- |
| IpfcXToolkitObsolete   | PRO_TK_OBSOLETE_FUNC | -34 |
Func
| IpfcXToolkitNoCoordSys | PRO_TK_NO_COORD_SYSTEM |     |
| ---------------------- | ---------------------- | --- |
-35
tem
| IpfcXToolkitAmbiguous | PRO_TK_E_AMBIGUOUS | -36 |
| --------------------- | ------------------ | --- |
| IpfcXToolkitDeadLock  | PRO_TK_E_DEADLOCK  | -37 |
| IpfcXToolkitBusy      | PRO_TK_E_BUSY      |     |
-38
| IpfcXToolkitInUse       | PRO_TK_E_IN_USE         | -39 |
| ----------------------- | ----------------------- | --- |
| IpfcXToolkitNoLicense   | PRO_TK_NO_LICENSE       | -40 |
| IpfcXToolkitBsplUnsuita | PRO_TK_BSPL_UNSUITABLE_ | -41 |
| bleDegree               | DEGREE                  |     |
| IpfcXToolkitBsplNonSt   | PRO_TK_BSPL_NON_STD_    |     |
-42
| dEndKnots               | END_KNOTS          |     |
| ----------------------- | ------------------ | --- |
| IpfcXToolkitBsplMultiIn | PRO_TK_BSPL_MULTI_ | -43 |
INNER_KNOTS
nerKnots
| IpfcXToolkitBadSrfCrv | PRO_TK_BAD_SRF_CRV    | -44 |
| --------------------- | --------------------- | --- |
| IpfcXToolkitEmpty     | PRO_TK_EMPTY          | -45 |
| IpfcXToolkitBadDimAt  | PRO_TK_BAD_DIM_ATTACH | -46 |
tach
| IpfcXToolkitNotDis | PRO_TK_NOT_DISPLAYED | -47 |
| ------------------ | -------------------- | --- |
played
| IpfcXToolkitCantModify  | PRO_TK_CANT_MODIFY   | -48 |
| ----------------------- | -------------------- | --- |
| IpfcXToolkitCheckoutCon | PRO_TK_CHECKOUT_     | -49 |
| flict                   | CONFLICT             |     |
| IpfcXToolkitCreateView  | PRO_TK_CRE_VIEW_BAD_ | -50 |
| BadSheet                | SHEET                |     |
| IpfcXToolkitCreateView  | PRO_TK_CRE_VIEW_BAD_ | -51 |
| BadModel                | MODEL                |     |
| IpfcXToolkitCreateView  | PRO_TK_CRE_VIEW_BAD_ | -52 |
| BadParent               | PARENT               |     |
OverviewoftheVBAPI 29

| IpfcXToolkitErrorChild | CreoTOOLKITError | #   |
| ---------------------- | ---------------- | --- |
Class
| IpfcXToolkitCreateView | PRO_TK_CRE_VIEW_BAD_ | -53 |
| ---------------------- | -------------------- | --- |
TYPE
BadType
| IpfcXToolkitCreateView | PRO_TK_CRE_VIEW_BAD_    | -54 |
| ---------------------- | ----------------------- | --- |
| BadExplode             | EXPLODE                 |     |
| IpfcXToolkitUnattached | PRO_TK_UNATTACHED_FEATS | -55 |
Feats
PRO_TK_REGEN_AGAIN
| IpfcXToolkitRegenerateA |     | -56 |
| ----------------------- | --- | --- |
gain
| IpfcXToolkitDrawingCrea | PRO_TK_DWGCREATE_ERRORS | -57 |
| ----------------------- | ----------------------- | --- |
teErrors
| IpfcXToolkitUnsupported | PRO_TK_UNSUPPORTED   | -58 |
| ----------------------- | -------------------- | --- |
| IpfcXToolkitNoPermis    | PRO_TK_NO_PERMISSION | -59 |
sion
| IpfcXToolkitAuthentica  | PRO_TK_AUTHENTICATION_ | -60 |
| ----------------------- | ---------------------- | --- |
| tionFailure             | FAILURE                |     |
| IpfcXToolkitMultibodyUn | PRO_TK_MULTIBODY_      | -69 |
| supported               | UNSUPPORTED            |     |
| IpfcXToolkitAppNoLi     | PRO_TK_APP_NO_LICENSE  |     |
-92
cense
| IpfcXToolkitAppExcess | PRO_TK_APP_XS_CALLBACKS | -93 |
| --------------------- | ----------------------- | --- |
Callbacks
| IpfcXToolkitAppStartup | PRO_TK_APP_STARTUP_FAIL |     |
| ---------------------- | ----------------------- | --- |
-94
Failed
| IpfcXToolkitAppInitiali | PRO_TK_APP_INIT_FAIL | -95 |
| ----------------------- | -------------------- | --- |
zationFailed
| IpfcXToolkitAppVersion | PRO_TK_APP_VERSION_ |     |
| ---------------------- | ------------------- | --- |
-96
| Mismatch               | MISMATCH                |     |
| ---------------------- | ----------------------- | --- |
| IpfcXToolkitAppCommuni | PRO_TK_APP_COMM_FAILURE | -97 |
cationFailure
PRO_TK_APP_NEW_VERSION
| IpfcXToolkitAppNewVer |     | -98 |
| --------------------- | --- | --- |
sion
Theexception IpfcXProdevErrorrepresentsa generalerrorthatoccurred
| whileexecutingaPro/DEVELOP | function andisequivalentto | an  |
| -------------------------- | -------------------------- | --- |
IpfcXToolkitGeneralErrorexception.(PTC doesnot recommendtheuse
ofPro/DEVELOP functions.)
TheIpfcXExternalDataError exceptiontypesanditschildren arethrown
fromExternalDatamethods.Seethechapter onExternalDataon page335for
moreinformation.
| 30  |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | ------------------------------- |

3
VB API Fundamentals:Controlling
Creo Parametric
Overview..................................................................................................................32
SimpleAsynchronousMode.......................................................................................32
StartingandStoppingCreoParametric ......................................................................32
ConnectingtoaCreoParametricProcess...................................................................35
FullAsynchronousMode............................................................................................37
TroubleshootingVBAPIApplications..........................................................................39
Thischapterexplains howtousethe VBAPI toestablish aconnectionto Creo
Parametric.
31

Overview
Asynchronousmodeisamultiprocess modein whichtheVB APIapplicationand
CreoParametriccan performconcurrent operations.TheVB APIapplication
(containingitsownmain()method) isstartedindependentlyof CreoParametric
andsubsequentlyeitherstartsor connectsto aCreo Parametricprocess.
Dependingon howyourasynchronous applicationhandlesmessagesfrom Creo
Parametric,your applicationcanbe classifiedas eithersimpleor full.The
followingsectionsdescribe simpleandfullasynchronousmode.
Simple Asynchronous Mode
Asimpleasynchronous applicationdoes notimplementa wayto handle
requestsfrom CreoParametric.Therefore, theVB APIcannotplant listenersto be
notifiedwhen eventshappen inCreo Parametric.Consequently, CreoParametric
cannotinvoke themethods thatmustbe suppliedwhen youadd,forexample,
menubuttonsto Creo Parametric.
Despitethis limitation,a simpleasynchronousmodeapplicationcanbe usedto
automateprocesses in CreoParametric.The applicationmayeitherstartor
connectto anexisting CreoParametric session,andmayaccessCreo Parametric
ininteractiveorin anongraphical,noninteractivemode.WhenCreo Parametric
isrunning withgraphics,itisan interactiveprocessavailabletothe user.
WhenyoudesignaVB API applicationto runinsimpleasynchronous mode,keep
thefollowingpointsin mind:
• TheCreo Parametricprocessandthe applicationperformoperations
concurrently.
• Noneof theapplication’s listenermethods canbe invokedbyCreo Parametric.
Starting and Stopping Creo Parametric
Thefollowingmethodsareused tostartandstopCreo Parametricwhen usingthe
VBAPI applications.
MethodsIntroduced:
• CCpfcAsyncConnection.Start()
• IpfcAsyncConnection.End()
AVBapplicationcanspawn andconnectto aCreo Parametricprocesswith the
methodCCpfcAsyncConnection.Start().
Afterthis methodreturnstheasynchronous connectionobject,the VBAPI
applicationcancall theCreo Parametricprocessusingthe appropriateAPIs.Inthe
interactivemode,youcanalso accesstheCreo Parametricsessionwhenitis
running.
32 Creo®ParametricVBAPIUser’sGuide

Theasynchronousapplicationisnot terminatedwhenCreo Parametricterminates.
Thisisuseful whentheapplicationneedsto performCreo Parametricoperations
intermittently, andtherefore,muststartandstopCreo Parametricmorethanonce
duringasession.
Theapplicationcanconnectto or startonly oneCreo Parametricsessionatany
time.IftheVB APIapplicationspawns asecondsession, connectionto thefirst
| sessionis | lost. |     |     |     |     |     |     |     |
| --------- | ----- | --- | --- | --- | --- | --- | --- | --- |
Toendany CreoParametricprocessthatthe applicationis connectedto, callthe
methodIpfcAsyncConnection.End().
| Setting |     | Up a Noninteractive |     |     | Session |     |     |     |
| ------- | --- | ------------------- | --- | --- | ------- | --- | --- | --- |
You canspawna CreoParametricsessionthatis bothnoninteractiveand
nongraphical.Inasynchronousmode,include thefollowingstringsin theCreo
Parametricstartor connectcallto CCpfcAsyncConnection.Start():
• -g:no_graphics—Turnoffthegraphicsdisplay.
| •   | -i:rpc_input—CausesCreo |     |     | Parametricto |     | expectinputfrom |     | your |
| --- | ----------------------- | --- | --- | ------------ | --- | --------------- | --- | ---- |
asynchronousapplicationonly.
Note
|     | Bothoftheseargumentsare |     |     | required,but |     | theorder | isnotimportant. |     |
| --- | ----------------------- | --- | --- | ------------ | --- | -------- | --------------- | --- |
Thesyntaxof thecallfor anoninteractive,nongraphicalsessionis asfollows:
| Dim | aC as | IpfcAsyncConnection         |     |     |     |     |     |     |
| --- | ----- | --------------------------- | --- | --- | --- | --- | --- | --- |
| Dim | ccAC  | as New CcpfcAsyncConnection |     |     |     |     |     |     |
aC = ccAC.Start ("pro -g:no_graphics -i:rpc_input",<text_dir>);
| wherepro |     | isthecommand | tostartCreo |           | Parametric. |     |     |     |
| -------- | --- | ------------ | ----------- | --------- | ----------- | --- | --- | --- |
| Example  |     | Code for     | Visual      | Basic.NET |             |     |     |     |
Thesamplecodein thefilepfcAsynchronousModeExamples.vb located
| at<creo_vbapi_loadpoint>/vbapi_examples |     |     |     |     |     |     | demonstrates | howto |
| --------------------------------------- | --- | --- | --- | --- | --- | --- | ------------ | ----- |
usethe VBAPI to startCreo Parametricasynchronously, retrievea Sessionandto
| openamodelin |     | Creo Parametric. |        |       |     |     |              |     |
| ------------ | --- | ---------------- | ------ | ----- | --- | --- | ------------ | --- |
| Example      |     | Code for         | Visual | Basic |     | for | Applications |     |
Thisexampledemonstratesthe VBAPI syntaxin amacro writtenin VisualBasic
forApplications,for example,aswouldbe runby abutton in aMicrosoftWord
documentorMicrosoft Excelspreadsheet.This exampleisidenticalto the
| previousexample,exceptfor |     |     | thesyntax. |     |     |     |     |     |
| ------------------------- | --- | --- | ---------- | --- | --- | --- | --- | --- |
VBAPIFundamentals:ControllingCreoParametric 33

| Private Sub         | btnRun_Click() |                      |                     |                      |     |     |     |
| ------------------- | -------------- | -------------------- | ------------------- | -------------------- | --- | --- | --- |
| Dim asyncConnection |                |                      | As                  | IpfcAsyncConnection  |     |     |     |
| Dim cAC             | As             | CCpfcAsyncConnection |                     |                      |     |     |     |
| Dim session         |                | As IpfcBaseSession   |                     |                      |     |     |     |
| Dim descModel       |                | As                   | IpfcModelDescriptor |                      |     |     |     |
| Dim descModelCreate |                |                      | As                  | CCpfcModelDescriptor |     |     |     |
| Dim model           | As             | IpfcModel            |                     |                      |     |     |     |
| Dim workDir         |                | As String            |                     |                      |     |     |     |
| Dim position        |                | As Integer           |                     |                      |     |     |     |
| On Error            | GoTo           | RunError             |                     |                      |     |     |     |
'================================================================-
======
First Argument : The path to the Creo Parametric executable along
with command
line options. -i and -g flags make Creo Parametric run in non-
graphic,
| 'non-interactive  |     | mode   |      |         |     |         |        |
| ----------------- | --- | ------ | ---- | ------- | --- | ------- | ------ |
| 'Second Argument: |     | String | path | to menu | and | message | files. |
'================================================================-
======
| Set cAC             | = New | CCpfcAsyncConnection |                             |     |     |     |            |
| ------------------- | ----- | -------------------- | --------------------------- | --- | --- | --- | ---------- |
| Set asyncConnection |       |                      | = cAC.Start(txtExePath.Text |     |     |     | + " -g:no_ |
graphics
| -i:rpc_input", |     |                           | ".") |     |     |     |     |
| -------------- | --- | ------------------------- | ---- | --- | --- | --- | --- |
| Set session    |     | = asyncConnection.session |      |     |     |     |     |
'================================================================-
======
| 'Get current | directory |           |     |     |     |     |     |
| ------------ | --------- | --------- | --- | --- | --- | --- | --- |
| 'Set it as   | working   | directory |     |     |     |     |     |
'================================================================-
======
| workDir                 | = ActiveWorkbook.FullName |                   |     |           |     |     |     |
| ----------------------- | ------------------------- | ----------------- | --- | --------- | --- | --- | --- |
| position                | =                         | InStrRev(workDir, |     | "\")      |     |     |     |
| workDir                 | = Left(workDir,           |                   |     | position) |     |     |     |
| session.ChangeDirectory |                           |                   |     | (workDir) |     |     |     |
'================================================================-
======
| 'VB api process |     | calls | and other | processing |     | to be | done |
| --------------- | --- | ----- | --------- | ---------- | --- | ----- | ---- |
'================================================================-
======
| Set descModelCreate |     |     | = New | CCpfcModelDescriptor |     |     |     |
| ------------------- | --- | --- | ----- | -------------------- | --- | --- | --- |
Set descModel = descModelCreate.Create(EpfcModelType.EpfcMDL_
PART,
|           |     | "partModel.prt",                 |     |     | dbnull) |     |     |
| --------- | --- | -------------------------------- | --- | --- | ------- | --- | --- |
| Set model | =   | session.RetrieveModel(descModel) |     |     |         |     |     |
'================================================================-
======
| End the Creo | Parametric |     | session | when | done |     |     |
| ------------ | ---------- | --- | ------- | ---- | ---- | --- | --- |
'================================================================-
======
| 34  |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

| If Not | asyncConnection           |     | Is Nothing |     | Then |     |     |     |
| ------ | ------------------------- | --- | ---------- | --- | ---- | --- | --- | --- |
| If     | asyncConnection.IsRunning |     |            |     | Then |     |     |     |
asyncConnection.End
End If
| End If |     |     |     |     |     |     |     |     |
| ------ | --- | --- | --- | --- | --- | --- | --- | --- |
RunError:
| If Err.Number | <>  | 0 Then |     |     |     |     |     |     |
| ------------- | --- | ------ | --- | --- | --- | --- | --- | --- |
MsgBox "Process Failed : Unknown error occured." + Chr(13) + _
|        | "Error                    | No: | " + CStr(Err.Number) |     |      | + Chr(13)   | +       | _   |
| ------ | ------------------------- | --- | -------------------- | --- | ---- | ----------- | ------- | --- |
|        | "Error:                   |     | " + Err.Description, |     |      | vbCritical, | "Error" |     |
| If Not | asyncConnection           |     | Is Nothing           |     | Then |             |         |     |
| If     | asyncConnection.IsRunning |     |                      |     | Then |             |         |     |
asyncConnection.End
End If
| End If |     |     |     |     |     |     |     |     |
| ------ | --- | --- | --- | --- | --- | --- | --- | --- |
End If
End Sub
Thefollowingfiguredisplaysthebutton in MicrosoftExceldesignedfor the
aboveapplication.
| Connecting |     | to  | a Creo | Parametric |     |     | Process |     |
| ---------- | --- | --- | ------ | ---------- | --- | --- | ------- | --- |
MethodsIntroduced:
• CCpfcAsyncConnection.Connect()
• CCpfcAsyncConnection.GetActiveConnection()
• IpfcAsyncConnection.Disconnect()
Asimpleasynchronousapplicationcan alsoconnect toa CreoParametricprocess
| thatis alreadyrunning          |     | ona | localcomputer.The |     | method                  |     |     |      |
| ------------------------------ | --- | --- | ----------------- | --- | ----------------------- | --- | --- | ---- |
| CCpfcAsyncConnection.Connect() |     |     |                   |     | performsthisconnection. |     |     | This |
methodfailstoconnect ifmultiple CreoParametricsessionsare running.If
severalversions ofCreo Parametricarerunning onthe samecomputer, tryto
connectby specifyinguseranddisplayparameters.However, ifseveralversions
ofCreo Parametricarerunningin thesameuser anddisplayparameters,the
| connectionmaynotbe |     | possible. |     |     |     |     |     |     |
| ------------------ | --- | --------- | --- | --- | --- | --- | --- | --- |
VBAPIFundamentals:ControllingCreoParametric 35

CCpfcAsyncConnection.GetActiveConnection() returnsthe current
| connectionto | aCreo Parametricsession.            |     |     |        |
| ------------ | ----------------------------------- | --- | --- | ------ |
| Todisconnect | froma CreoParametricprocess,callthe |     |     | method |
IpfcAsyncConnection.Disconnect().
| Connecting | Via | Connection | ID  |     |
| ---------- | --- | ---------- | --- | --- |
MethodsIntroduced:
• IpfcAsyncConnection.GetConnectionId()
• IpfcConnectionId.ExternalRep
• CCpfcConnectionId.Create()
• IpfcAsyncConnection.ConnectById()
EachCreoParametricprocessmaintainsaunique identityfor communications
| purposes.Usethis | IDto | reconnectto | aCreo Parametricprocess. |     |
| ---------------- | ---- | ----------- | ------------------------ | --- |
ThemethodIpfcAsyncConnection.GetConnectionId() returnsa data
structurecontainingtheconnectionID.
Iftheconnectionid mustbe passedto someotherapplicationthe method
IpfcConnectionId.ExternalRepprovidesthe
stringexternal
| representationforthe |     | connectionID. |     |     |
| -------------------- | --- | ------------- | --- | --- |
ThemethodCCpfcConnectionId.Create() takesa stringrepresentation
andcreatesaConnectionIddataobject.Themethod
IpfcAsyncConnection.ConnectById()connects toCreo Parametricat
thespecifiedconnectionID.
Note
ConnectionIDsareunique for eachCreoParametricprocessand arenot
| maintainedafteryouquitCreo |           |            | Parametric. |     |
| -------------------------- | --------- | ---------- | ----------- | --- |
| Status                     | of a Creo | Parametric | Process     |     |
MethodIntroduced:
• IpfcAsyncConnection.IsRunning()
Tofind outwhetheraCreo Parametricprocessisrunning, usethemethod
IpfcpfcAsyncConnection.IsRunning().
| 36  |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

| Getting | the | Session |     | Object |     |     |     |     |
| ------- | --- | ------- | --- | ------ | --- | --- | --- | --- |
MethodIntroduced:
• IpfcAsyncConnection.Session()
ThemethodIpfcAsyncConnection.Session() returnsthesessionobject
representingthe CreoParametricsession. Use thisobjectto accessthecontents of
theCreoParametricsession. Seethe SessionObjectsonpage43 chapterfor
additionalinformation.
| Full | Asynchronous |     |     |     | Mode |     |     |     |
| ---- | ------------ | --- | --- | --- | ---- | --- | --- | --- |
Fullasynchronousmode isidenticalto thesimpleasynchronousmodeexceptin
thewaytheVB APIapplicationhandlesrequests fromCreo Parametric.Insimple
| asynchronousmode,itisnotpossibleto |     |     |     |     | processtheserequests. |     | Infull |     |
| ---------------------------------- | --- | --- | --- | --- | --------------------- | --- | ------ | --- |
asynchronousmode,theapplicationimplementsa controlloopthat‘‘listens’’for
messagesfromCreo Parametric.Asaresult,Creo Parametriccancallfunctions in
theapplication,includingcallback functionsfor menubuttons andnotifications.
Note
| Using           | fullasynchronousmode |                     |     | requiresstartingorconnectingto |                            |     |     | Creo |
| --------------- | -------------------- | ------------------- | --- | ------------------------------ | -------------------------- | --- | --- | ---- |
| Parametricusing |                      | themethodsdescribed |     |                                | in theprevioussections.The |     |     |      |
differenceisthatthe applicationmustprovidean eventloop to processcalls
| from | menubuttons | andlisteners. |     |     |     |     |     |     |
| ---- | ----------- | ------------- | --- | --- | --- | --- | --- | --- |
MethodsIntroduced:
• IpfcAsyncConnection.EventProcess()
• IpfcAsyncConnection.WaitForEvents()
• IpfcAsyncConnection.InterruptEventProcessing()
• IpfcAsyncActionListener.OnTerminate()
Thecontrolloop ofan applicationrunning in full asynchronousmodemust
| containacallto | themethod |     | IpfcAsyncConnection.EventProcess(), |     |     |     |     |     |
| -------------- | --------- | --- | ----------------------------------- | --- | --- | --- | --- | --- |
whichtakesnoarguments. Thismethod allowstheapplicationto respondto
messagessentfrom CreoParametric.For example,if theuserselectsamenu
| buttonthatisadded                  |     | byyour | application, |     |     |           |           |      |
| ---------------------------------- | --- | ------ | ------------ | --- | --- | --------- | --------- | ---- |
| IpfcAsyncConnection.EventProcess() |     |        |              |     |     | processes | thecallto | your |
listenerandreturnswhen thecallcompletes. Formoreinformation onlisteners
| andaddingmenu                                |     | buttons,seetheSession |     |     | Objectson | page43chapter. |            |     |
| -------------------------------------------- | --- | --------------------- | --- | --- | --------- | -------------- | ---------- | --- |
| ThemethodIpfcAsyncConnection.WaitForEvents() |     |                       |     |     |           |                | providesan |     |
alternativeto thedevelopmentofan event processingloop in afullasynchronous
modeapplication.Callthis functionto havetheapplicationwaitin aloop for
VBAPIFundamentals:ControllingCreoParametric 37

eventsto bepassedfrom Creo Parametric.No otherprocessingtakesplace while
| theapplicationis                                 | waiting.Theloop |     | continuesuntil |     |     |     |        |
| ------------------------------------------------ | --------------- | --- | -------------- | --- | --- | --- | ------ |
| IpfcAsyncConnection.InterruptEventProcessing()is |                 |     |                |     |     |     | called |
fromaVB callbackaction,oruntil theapplicationdetectstheterminationofCreo
Parametric.
Itisoftennecessary foryour fullasynchronous applicationto be notifiedof the
terminationof theCreo Parametricprocess.Inparticular, yourcontrolloop need
notcontinuetolisten forCreo Parametricmessagesif CreoParametricis no
longerrunning.
AnAsyncConnection objectcanbe assignedan ActionListener to binda
terminationactionthatis executedupontheterminationof CreoParametric.The
| methodIpfcAsyncActionListener.OnTerminate() |     |     |                    |              |         |            | handlesthe |
| ------------------------------------------- | --- | --- | ------------------ | ------------ | ------- | ---------- | ---------- |
| terminationthatyoumustoverride.It           |     |     |                    | sendsamember |         | ofthe      | class      |
| IpfcTerminationStatus,whichis               |     |     |                    | oneofthe     |         | following: |            |
| • EpfcTERM_EXIT—Normal                      |     |     | exit(theuserclicks |              |         | Exitonthe  | menu).     |
| • EpfcTERM_ABNORMAL—Quit                    |     |     | witherror          |              | status. |            |            |
| • EpfcTERM_SIGNAL—Fatalsignal               |     |     |                    | raised.      |         |            |            |
Yourapplicationcan interprettheterminationtypeandtakeappropriateaction.
Formoreinformationon Action Listeners,seetheAction Listenersonpage267
chapter.
| Example | Code |     |     |     |     |     |     |
| ------- | ---- | --- | --- | --- | --- | --- | --- |
Thesamplecodein thefilepfcAsynchronousModeExamples.vb located
at<creo_vbapi_loadpoint>/vbapi_examples is afullyasynchronous
application.Itfollows theprocedurefora fullasynchronous application:
1. Theapplicationestablisheslisteners forCreo Parametricevents,in thiscase,
| themenubutton | andtheterminationlistener. |     |     |     |     |     |     |
| ------------- | -------------------------- | --- | --- | --- | --- | --- | --- |
2. Theapplicationgoesintoacontrolloop calling EventProcesswhichallows
| theapplicationtorespond |          | to  | theCreo | Parametricevents. |     |     |     |
| ----------------------- | -------- | --- | ------- | ----------------- | --- | --- | --- |
| Messageand              | MenuFile |     |         |                   |     |     |     |
#
#
VB-Async
VB-Async
#
#
USER#Async#App
Async Button
#
#
USER#Async#Help
| 38  |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

| Button added | via Async | Application |     |     |
| ------------ | --------- | ----------- | --- | --- |
#
#
| Troubleshooting      |          | VB                      | API Applications |     |
| -------------------- | -------- | ----------------------- | ---------------- | --- |
| General              | Problems |                         |                  |     |
| IpfcXToolkitNotFound |          | exceptiononthefirstcall |                  | to  |
CCpfcAsyncConnection.StartonWindows
Makesureyour commandiscorrect.Ifitisnot afullpathto ascript or
executable,make sure$PATHis setcorrectly.Try fullpathin thecommand:if it
| works,thenyour$PATH |     | isincorrect. |     |     |
| ------------------- | --- | ------------ | --- | --- |
IpfcXToolkitGeneralErroror IpfcXToolkitCommError onthefirstcall to
CCpfcAsyncConnection.StartorCCpfcAsyncConnection.Connect
• Makesuretheenvironment variablePRO_COMM_MSG_EXE issetto thefull
| pathtopro_comm_msg, |     | including | <filename.exe>. |     |
| ------------------- | --- | --------- | --------------- | --- |
• Makesuretheenvironment variablePRO_DIRECTORY is setto the Creo
Parametricinstallationdirectory.
| • Makesurenameservice(nmsd)is |     |     | running. |     |
| ----------------------------- | --- | --- | -------- | --- |
CCpfcAsyncConnection.Starthangs,eventhoughCreo Parametric
alreadystarted
Makesurename service(nmsd)isalso startedalongwithCreo Parametric.Open
| TaskManager | andlookfor | nmsd.exein | theprocesslisting. |     |
| ----------- | ---------- | ---------- | ------------------ | --- |
VBAPIFundamentals:ControllingCreoParametric 39

4
The VB API Online Browser
OnlineDocumentationVBAPIWizard.........................................................................41
Thischapterdescribes howtousetheonlinebrowserprovided withthe VB
APIWizard.
40 Creo®ParametricVBAPIUser’sGuide

| Online | Documentation |     |     | VB APIWizard |     |
| ------ | ------------- | --- | --- | ------------ | --- |
TheVB APIprovidesan onlinebrowsercalledtheVB APIWizardthatdisplays
detaileddocumentation.This browserdisplaysinformationfrom theVBAPI
User’s GuideandAPI specificationsderivedfromthe VBAPI headerfiledata.
| TheVB           | APIWizardcontains | thefollowingitems: |     |     |     |
| --------------- | ----------------- | ------------------ | --- | --- | --- |
| • Definitionsof | theVB             | APImodules         |     |     |     |
• Definitionsof theVB APIclasses andinterfacesandtheirhierarchical
relationships
| • Descriptionsof |           | theVB APImethods      |         |              |       |
| ---------------- | --------- | --------------------- | ------- | ------------ | ----- |
| • Declarationsof |           | datatypesused         | bytheVB | API methods  |       |
| • TheVB          | APIUser’s | Guidethatyoucanbrowse |         | bytopic orby | class |
• Codeexamplesforthe VBAPI methods(taken fromsampleapplications
| providedaspart |     | ofthe VBAPI | installation) |     |     |
| -------------- | --- | ----------- | ------------- | --- | --- |
Readthe ReleaseNotes andREADMEfileforthe mostup-to-dateinformationon
documentationchanges.
Note
• TheVB APIUser’s Guideisalsoavailablein PDFformatatthefollowing
location:
<creo_vbapi_loadpoint>\vbug.pdf
• FromCreo 4.0F000,the appletbasedAPIWizardisno longersupported. Use
| thenon-applet | basedAPIWizardinstead. |     |     |     |     |
| ------------- | ---------------------- | --- | --- | --- | --- |
| APIWizard     | Overview               |     |     |     |     |
TheAPIWizardsupports InternetExplorer, Firefox,andChromium browsers.
| Startthe | VBAPIWizardbypointing |     | yourbrowserto: |     |     |
| -------- | --------------------- | --- | -------------- | --- | --- |
<creo_vbapi_loadpoint>\vbapidoc\index.html
Apagecontaininglinksto theVB APIWizardandUser’s Guidewillopen inthe
webbrowser.
TheVBAPIOnlineBrowser 41

| Non-Applet | APIWizard | Top Page |
| ---------- | --------- | -------- |
Thetoppageofnon-applet basedAPIWizardhaslinksto theVB APIAPIWizard
andUser’sGuide.The APIWizardopens an HTMLpagethatcontains linksto VB
APIclassesandrelated functions.The User’sGuideopens an HTMLpagethat
displaysthe Tableof Contentsof theUser’s Guide,withlinksto thechapters,and
| sectionsunderthe | chapters. |     |
| ---------------- | --------- | --- |
ClickAPIWizardto open thelist ofVB APIclasses andrelatedfunctions.Click a
| classor functionnameto | read | moreaboutit. |
| ---------------------- | ---- | ------------ |
You cansearchfor specifiedinformationin theAPIWizard.Use thesearchfield at
thetopleftpanetosearch forfunctions.Youcan searchfor informationusing the
followingcriteria:
| • Searchby | functionnames |     |
| ---------- | ------------- | --- |
• Searchusing wildcardcharacter *,where*(asterisk) matcheszero ormore
nonwhitespacecharacters
Theresulting functionnames aredisplayedina dropdownlist with linksto html
pages.
You canalsohoverthe mouseover afteryouentera stringinthe searchfield.
| Thefollowingsearchoptions     |     | aredisplayed:        |
| ----------------------------- | --- | -------------------- |
| • Class/Functions—Searchesfor |     | classesandfunctions. |
| • Exceptions—Searchesonly     |     | for exceptions.      |
| • Enumeration—Searchesonlyfor |     | enumerations.        |
Selectan option andthesearch resultsaredisplayedbasedonthis criteria.
| User’s      | Guide               |                 |
| ----------- | ------------------- | --------------- |
| ClickUser’s | Guideto accesstheVB | APIUser’sGuide. |
42 Creo®ParametricVBAPIUser’sGuide

5
Session Objects
OverviewofSessionObjects......................................................................................44
CompatibilityofDeprecatedMethods..........................................................................44
Directories................................................................................................................45
AccessingtheCreoParametricInterface.....................................................................48
Thischapterdescribes howtoprogramon thesessionlevel usingthe VBAPI.
43

| Overview | of Session | Objects |     |
| -------- | ---------- | ------- | --- |
TheCreo ParametricSessionobject(containedin theclass IpfcSession) is
thehighestlevelobjectin theVB API.Anyprogramthataccessesdatafrom Creo
Parametricmustfirst getahandleto theIpfcSession objectbeforeaccessing
morespecificdata.
TheIpfcSession objectcontainsmethods toperformthefollowingoperations:
• Accessingmodelsandwindows (describedin the ModelsandWindows
chapters).
| • Working withthe                     | CreoParametricuserinterface. |        |             |
| ------------------------------------- | ---------------------------- | ------ | ----------- |
| • Allowinginteractiveselectionofitems |                              | within | thesession. |
• Accessingglobalsettingssuchas linestyles,colors,andconfigurationoptions.
Thefollowingsectionsdescribetheseoperationsin detail.Referto thechapter
"ControllingCreoParametric"for moreinformationon howto connectto aCreo
Parametricsession.
| Compatibility | of Deprecated |     | Methods |
| ------------- | ------------- | --- | ------- |
Inareleasecycle,somemethods aredeprecated,andnewmethods areadded.The
deprecatedmethods aresupportedin thecurrentrelease,andthenobsoletedina
futurerelease.You caneitherchooseto letthedeprecatedmethods workinthe
currentrelease,orallow onlynewmethods towork. Use themethodsexplained in
thissectionalongwith thecompatibilityvalue towork witheitherdeprecatedor
| newmethodsforthe | currentrelease. |     |     |
| ---------------- | --------------- | --- | --- |
MethodsIntroduced:
• IpfcAppInfo.GetCompatibility()
• IpfcAppInfo.SetCompatibility()
• CCpfcAppInfo.Create()
• IpfcSession.GetAppInfo()
• IpfcSession.SetAppInfo()
ThemethodsIpfcAppInfo.GetCompatibility() and
IpfcAppInfo.SetCompatibility()
getandsetthecompatibility value
| forthespecifiedapplicationusing |     | theenumerateddatatype |      |
| ------------------------------- | --- | --------------------- | ---- |
| EpfcCreoCompatibility.The       |     | validvalues           | are: |
• EpfcCompatibilityUndefined—Specifies thatcompatibility valueis
| notset. Thedefaultcompatibility |     | valueisused. |     |
| ------------------------------- | --- | ------------ | --- |
• EpfcC3Compatible—Specifiesthatthemethods deprecatedinthe VB
API4.0 arecompatibleandcontinueworking in theVB API4.0.Bydefault
| thecompatibilityis | settoEpfcC3Compatible. |     |                                 |
| ------------------ | ---------------------- | --- | ------------------------------- |
| 44                 |                        |     | Creo®ParametricVBAPIUser’sGuide |

• EpfcC4Compatible—Specifiesthatthemethods deprecatedinthe VB
| API4.0 | willnot workinthe | VBAPI | 4.0.Ifyourapplicationusesthe |     |
| ------ | ----------------- | ----- | ---------------------------- | --- |
deprecatedmethods,youmustreplace thesemethodswith newmethodsand
rebuildyouapplications.
Note
Ifyourebuildor runthe VBAPI applicationsfrompreviousreleasesin the
currentrelease,thecompatibilityissetEpfcC3Compatible to current
release.Forexample,if yourebuildatheVB API3.0 applicationin release
4.0,thecompatibility issetto C3Compatible. Tosetthe compatibilityto
thecurrentrelease,usethemethod
IpfcAppInfo.SetCompatibility().
| ThemethodCCpfcAppInfo.Create() |     |     | createsa newinstanceof | the |
| ------------------------------ | --- | --- | ---------------------- | --- |
pfcAppInfoobject.
| ThemethodsIpfcSession.GetAppInfo() |     |     | and |     |
| ---------------------------------- | --- | --- | --- | --- |
IpfcSession.SetAppInfo()get andsettheinformationfor an application
intermsof theircompatibilityvalue asapfcSession.AppInfo object.
Directories
MethodsIntroduced:
• IpfcBaseSession.GetCurrentDirectory()
• IpfcBaseSession.ChangeDirectory()
| ThemethodIpfcBaseSession.GetCurrentDirectory() |     |     |     | returnsthe |
| ---------------------------------------------- | --- | --- | --- | ---------- |
absolutepathnamefor thecurrentworkingdirectoryof CreoParametric.
| ThemethodIpfcBaseSession.ChangeDirectory() |     |            | changesCreo |     |
| ------------------------------------------ | --- | ---------- | ----------- | --- |
| Parametrictoanotherworking                 |     | directory. |             |     |
File Handling
MethodsIntroduced:
• IpfcBaseSession.ListFiles()
• IpfcBaseSession.ListSubdirectories()
ThemethodIpfcBaseSession.ListFiles() returnsalist offilesin a
directory, giventhedirectorypath. You canfilter thelist toinclude onlyfilesof a
| particulartype, | asspecifiedbythe | fileextension. |     |     |
| --------------- | ---------------- | -------------- | --- | --- |
SessionObjects 45

| Startingwith                       | Pro/ENGINEERWildfire5.0 |     |     | M040,themethod |              |             |
| ---------------------------------- | ----------------------- | --- | --- | -------------- | ------------ | ----------- |
| IpfcBaseSession.ListFiles()canalso |                         |     |     |                | listinstance | objectswhen |
accessingWindchillworkspacesorfolders.APDM location(forworkspaceor
commonspace)mustbepassedas thedirectorypath.The followingoptionshave
| beenadded | inthe | EpfcFileListOptenumerated |     |     | type: |     |
| --------- | ----- | ------------------------- | --- | --- | ----- | --- |
• EpfcFILE_LIST_ALL—Lists allthe files.It mayalso includemultiple
| versionsof | thesamefile. |     |     |     |     |     |
| ---------- | ------------ | --- | --- | --- | --- | --- |
• EpfcFILE_LIST_LATEST—Listsonly thelatestversion ofeachfile.
| EpfcFILE_LIST_ALL_INST—Same                   |                          |                  |         |                | EpfcFILE_LIST_ALL   |            |
| --------------------------------------------- | ------------------------ | ---------------- | ------- | -------------- | ------------------- | ---------- |
| •                                             |                          |                  |         | asthe          |                     |            |
| option.It                                     | returnsinstances         |                  | onlyfor | PDM locations. |                     |            |
| • EpfcFILE_LIST_LATEST_INST—Same              |                          |                  |         |                | astheEpfcFILE_LIST_ |            |
| LATESToption.It                               |                          | returnsinstances |         | onlyfor        | PDM locations.      |            |
| ThemethodIpfcBaseSession.ListSubdirectories() |                          |                  |         |                |                     | returnsthe |
| subdirectoriesin                              | agivendirectorylocation. |                  |         |                |                     |            |
| Configuration                                 |                          | Options          |         |                |                     |            |
MethodsIntroduced:
• IpfcBaseSession.GetConfigOptionValues()
• IpfcBaseSession.SetConfigOption()
• IpfcBaseSession.LoadConfigFile()
You canaccessconfigurationoptions programmaticallyusing themethods
| describedin | thissection. |     |     |     |     |     |
| ----------- | ------------ | --- | --- | --- | --- | --- |
Usethe methodIpfcBaseSession.GetConfigOptionValues() to
retrievethevalue ofaspecifiedconfigurationfileoption.Pass theName ofthe
configurationfileoption astheinput tothis method.Themethodreturns an array
ofvaluesthatthe configurationfileoption issetto. Itreturnsa singlevalueif the
configurationfileoption isnota multi-valuedoption.The methodreturnsa nullif
| thespecifiedconfigurationfileoption        |     |     |     | doesnotexist. |     |                 |
| ------------------------------------------ | --- | --- | --- | ------------- | --- | --------------- |
| ThemethodIpfcBaseSession.SetConfigOption() |     |     |     |               |     | isused tosetthe |
valueofa specifiedconfigurationfileoption.Ifthe optionis amulti-valueoption,
| itadds anewvalueto                        |     | thearrayof | valuesthatalreadyexist. |     |     |                |
| ----------------------------------------- | --- | ---------- | ----------------------- | --- | --- | -------------- |
| ThemethodIpfcBaseSession.LoadConfigFile() |     |            |                         |     |     | loadsan entire |
configurationfileintoCreoParametric.
Macros
MethodIntroduced:
• IpfcBaseSession.RunMacro()
| 46  |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

| ThemethodIpfcBaseSession.RunMacro() |     |     | runsamacro | string.AVB |
| ----------------------------------- | --- | --- | ---------- | ---------- |
APImacro stringisequivalentto aCreo Parametricmapkeyminusthe key
sequenceandthe mapkeyname.To generatea macrostring,createamapkey in
CreoParametric.Referto theCreo Parametriconlinehelpfor moreinformation
aboutcreatingamapkey.
Copythe Valueofthe generatedmapkeyOption fromthe Tools▶▶Options dialog
| box.An | exampleValueis | asfollows: |     |     |
| ------ | -------------- | ---------- | --- | --- |
$F2 @MAPKEY_LABELtest;
| ~ Activate     | `main_dlg_cur` | `ProCmdModelNew.file`; |                       |     |
| -------------- | -------------- | ---------------------- | --------------------- | --- |
| ~ Activate     | `new` `OK`;    |                        |                       |     |
| Thekeysequence | is$F2.The      | mapkeynameis           | @MAPKEY_LABELtest.The |     |
remainderofthe stringfollowingthe firstsemicolon isthemacro stringthat
| shouldbepassedto | themethod       | IpfcBaseSession.RunMacro(). |     |     |
| ---------------- | --------------- | --------------------------- | --- | --- |
| Inthiscase,      | itis asfollows: |                             |     |     |
| ~ Activate       | `main_dlg_cur`  | `ProCmdModelNew.file`;      |     |     |
| ~ Activate       | `new` `OK`;     |                             |     |     |
Note
Creatingorediting themacro stringmanuallyisnot supportedasthe mapkeys
arenota supportedscripting language.The syntaxisnot definedforusersand
isnot guaranteedto remainconstant acrossdifferent datecodesofCreo
Parametric.
Macrosareexecutedfromsynchronous modeonly whencontrolreturns to Creo
Parametricfromthe VBAPI program.Macros arestoredin reverseorder (lastin,
firstout).
Macrosareexecutedassoonas theyare registered.Macrosare runin thesame
orderthattheyaresaved.
| Colors | and Line | Styles |     |     |
| ------ | -------- | ------ | --- | --- |
MethodsIntroduced:
• IpfcBaseSession.SetStdColorFromRGB()
• IpfcBaseSession.GetRGBFromStdColor()
• IpfcBaseSession.SetTextColor()
• IpfcBaseSession.SetLineStyle()
Thesemethodscontrolthegeneral displayof aCreo Parametricsession.
| Usethe         | methodIpfcBaseSession.SetStdColorFromRGB() |                           |     | to  |
| -------------- | ------------------------------------------ | ------------------------- | --- | --- |
| customizeanyof | theCreo                                    | Parametricstandardcolors. |     |     |
SessionObjects 47

| Tochangethecolorof |     | anytextin |     | thewindow, |     | usethemethod |     |
| ------------------ | --- | --------- | --- | ---------- | --- | ------------ | --- |
IpfcBaseSession.SetTextColor().
Tochangetheappearanceof nonsolidlines(forexample,datums)usethe method
IpfcBaseSession.SetLineStyle().
| Accessing |     | the | Creo |     | Parametric |     | Interface |
| --------- | --- | --- | ---- | --- | ---------- | --- | --------- |
TheIpfcSession objecthasmethodsthatworkwith theCreoParametric
| interface.These | methodsprovideaccessto |     |      |     |     | themessagewindow. |     |
| --------------- | ---------------------- | --- | ---- | --- | --- | ----------------- | --- |
| The Text        | Message                |     | File |     |     |                   |     |
Atextmessage fileis whereyou definestringsthataredisplayedin theCreo
Parametricuserinterface.This includes thestringson thecommandbuttons that
youaddtothe CreoParametricnumber,the helpstringthatdisplayswhenthe
user’scursor is positionedoversuch acommandbutton,and textstringsthatyou
displayin theMessage Window. You havetheoption ofincludinga translationfor
| eachstring   | inthe | textmessagefile. |      |         |     |      |     |
| ------------ | ----- | ---------------- | ---- | ------- | --- | ---- | --- |
| Restrictions | on    | the              | Text | Message |     | File |     |
You mustobserve thefollowingrestrictionswhenyouname yourmessagefile:
• Thename ofthefilemustbe 30characters orless, includingtheextension.
| • Thename                | ofthefilemustcontainlowercase |     |                  |           |     | charactersonly. |     |
| ------------------------ | ----------------------------- | --- | ---------------- | --------- | --- | --------------- | --- |
| • Thefileextensionmustbe |                               |     | threecharacters. |           |     |                 |     |
| • Theversion             | numbermustbein                |     |                  | therange1 |     | to9999.         |     |
• Allmessagefilenamesmustbe unique,andall messagekeystringsmustbe
uniqueacross all applicationsthatrunwith CreoParametric.Duplicate
messagefilenamesormessagekey stringscancauseCreo Parametricto
exhibitunexpectedbehavior.Toavoidconflictswith thenamesof Creo
Parametricorforeign applicationmessage filesor messagekeystrings,PTC
recommendsthatyouchooseaprefix uniqueto yourapplication, andprepend
| thatprefixto    | eachmessagefilename |                 |     |     | andeachmessagekeystring |     |     |
| --------------- | ------------------- | --------------- | --- | --- | ----------------------- | --- | --- |
| correspondingto |                     | thatapplication |     |     |                         |     |     |
Note
Message filesareloadedintoCreo Parametriconlyonceduringa session.If
youmakea changetothe messagefilewhile CreoParametricis runningyou
mustexitandrestart CreoParametricbeforethe changewilltakeeffect.
| 48  |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

| Contents |     | of the | Message |     | File |     |     |     |     |     |     |
| -------- | --- | ------ | ------- | --- | ---- | --- | --- | --- | --- | --- | --- |
Themessagefileconsistsof groups offourlines, onegroupfor eachmessageyou
| wantto | write. | Thefourlinesareas |     |     |     | follows: |     |     |     |     |     |
| ------ | ------ | ----------------- | --- | --- | --- | -------- | --- | --- | --- | --- | --- |
1. Astringthatactsas theidentifier forthe message.This keyword mustbe
|     | uniquefor           | allCreo |     | Parametricmessages. |     |                |     |     |     |     |     |
| --- | ------------------- | ------- | --- | ------------------- | --- | -------------- | --- | --- | --- | --- | --- |
| 2.  | Thestringthatwillbe |         |     | substitutedfor      |     | theidentifier. |     |     |     |     |     |
Thisstringcan includeplaceholdersfor run-timeinformationstoredin a
stringseqobject(shown in Writing Messagesto theMessage Window).
3. Thetranslationof themessageintoanotherlanguage(canbe blank).
| 4.      | Anintentionallyblanklinereservedfor |           |     |     |       | futureextensions. |     |        |     |        |     |
| ------- | ----------------------------------- | --------- | --- | --- | ----- | ----------------- | --- | ------ | --- | ------ | --- |
| Writing |                                     | a Message |     |     | Using | a Message         |     | Pop-up |     | Dialog |     |
Box
MethodIntroduced:
• IpfcSession.UIShowMessageDialog()
| ThemethodIpfcSession.UIShowMessageDialog() |                          |     |                  |     |     |                   |     |         | displaystheUI |     |     |
| ------------------------------------------ | ------------------------ | --- | ---------------- | --- | --- | ----------------- | --- | ------- | ------------- | --- | --- |
| messagedialog.The                          |                          |     | inputargumentsto |     |     | themethodare:     |     |         |               |     |     |
| •                                          | Message—Themessagetextto |     |                  |     |     | be displayedinthe |     | dialog. |               |     |     |
• Options—Aninstance oftheIpfcMessageDialogOptions containing
otheroptionsfor theresulting displayedmessage.Ifthisis notsupplied,the
dialogwillshowa defaultmessagedialogwithan Infoclassificationandan
OKbutton.If thisisnot tobe null, createan instance ofthis optionstypewith
|     | CCpfcMessageDialogOptions.Create(). |     |     |     |     |     |     | You | cansetthefollowing |     |     |
| --- | ----------------------------------- | --- | --- | --- | --- | --- | --- | --- | ------------------ | --- | --- |
options:
|     | ○ Buttons—Specifiesan            |     |     |     | arrayofbuttons |     | to  | includein         |     |     |     |
| --- | -------------------------------- | --- | --- | --- | -------------- | --- | --- | ----------------- | --- | --- | --- |
|     | IpfcMessageDialogOptions.Buttons |     |     |     |                |     |     | to setthisoption. |     |     |     |
DefaultButton—Specifies
|     | ○   |     |     |     |     | theidentifier |     | ofthedefault |     | buttonfor | the |
| --- | --- | --- | --- | --- | --- | ------------- | --- | ------------ | --- | --------- | --- |
dialogbox. Thismustmatchoneof theavailablebuttons.Use themethod
|     | IpfcMessageDialogOptions.DefaultButton |     |     |     |     |     |     |     | to  | setthisoption. |     |
| --- | -------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- | -------------- | --- |
○ DialogLabel—Thetextto displayasthe titleof thedialogbox. Ifnot
|     | supplied,thelabel                    |     |     | willbethe |     | EnglishstringInfo.Use |     |               | themethod         |     |     |
| --- | ------------------------------------ | --- | --- | --------- | --- | --------------------- | --- | ------------- | ----------------- | --- | --- |
|     | IpfcMessageDialogOptions.DialogLabel |     |     |           |     |                       |     |               | to setthisoption. |     |     |
|     | ○ MessageDialogType—The              |     |     |           |     | typeoficontobe        |     | displayedwith |                   |     | the |
dialogbox(Info, Prompt,Warning, orError). Ifnotsupplied,an Infoiconis
|                | used.Use                               |     | themethod |     |     |     |     |     |     |                |     |
| -------------- | -------------------------------------- | --- | --------- | --- | --- | --- | --- | --- | --- | -------------- | --- |
|                | MessageDialogOptions.MessageDialogType |     |           |     |     |     |     |     | to  | setthisoption. |     |
| SessionObjects |                                        |     |           |     |     |     |     |     |     |                | 49  |

| Accessing | the | Message | Window |     |
| --------- | --- | ------- | ------ | --- |
Thefollowingsectionsdescribehowto accessthemessagewindow usingthe VB
| API.Thetopics | areasfollows: |                   |          |           |
| ------------- | ------------- | ----------------- | -------- | --------- |
| • Writing     | Messagesto    | theMessage        | Windowon | page50    |
| • Writing     | Messagesto    | an InternalBuffer |          | on page50 |
| Writing       | Messages      | to the Message    |          | Window    |
MethodsIntroduced:
• IpfcSession.UIDisplayMessage()
• IpfcUIDisplayLocalizedMessage()
• IpfcSession.UIClearMessage()
Thesemethodsenableyouto displayprograminformationon thescreen.
Theinputargumentsto themethods IpfcSession.UIDisplayMessage()
andIpfcSession.UIDisplayLocalizedMessage() includethenames
ofthemessagefile, amessageidentifier, and(optionally)a stringseqobject
| thatcontains                       | upto 10pieces | ofrun-timeinformation.For |     |                           |
| ---------------------------------- | ------------- | ------------------------- | --- | ------------------------- |
| IpfcSession.UIDisplayMessage(),the |               |                           |     | stringsin thestringseqare |
identifiedas %0s,%1s,...%9sbasedontheir locationinthe sequence.For
IpfcSession.UIDisplayLocalizedMessage(),the stringsinthe
stringseqareidentifiedas%0w, %1w,...%9w basedontheirlocationin the
sequence.Toinclude othertypesofrun-timedata(suchas integersorreals)you
mustfirst convertthedatato stringsandstore itinthe stringsequence.
| Writing | Messages | to an Internal |     | Buffer |
| ------- | -------- | -------------- | --- | ------ |
MethodsIntroduced:
• IpfcBaseSession.GetMessageContents()
• IpfcBaseSession.GetLocalizedMessageContents()
ThemethodsIpfcBaseSession.GetMessageContents() and
IpfcBaseSession.GetLocalizedMessageContents() enableyouto
writeamessageto aninternalbuffer insteadofthe CreoParametricmessagearea.
Thesemethodstakethesameinput argumentsand performexactlythesame
| argumentsubstitutionand        |     | translationas | the |     |
| ------------------------------ | --- | ------------- | --- | --- |
| IpfcSession.UIDisplayMessage() |     |               |     | and |
IpfcSession.UIDisplayLocalizedMessage()methods describedin
theprevioussection.
| 50  |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

| Message | Classification |     |     |     |     |
| ------- | -------------- | --- | --- | --- | --- |
Messagesdisplayedin theVB APIinclude asymbol thatidentifiesthemessage
type.Everymessagetypeisidentifiedby aclassificationthatbegins withthe
characters%C.Amessageclassification requiresthatthe messagekeyline(line
| oneinthe | messagefile)mustbepreceded |     |     | by theclassificationcode. |     |
| -------- | -------------------------- | --- | --- | ------------------------- | --- |
Note
Anymessagekeystringusedin the codeshouldnot containtheclassification.
theVBAPI applicationscannowdisplayanyorall ofthe followingmessage
symbols:
• Prompt—This VBAPI messageispreceded bya greenarrow. Theusermust
| respondto | thismessagetype. |     | Respondingincludes,specifyinginput |     |     |
| --------- | ---------------- | --- | ---------------------------------- | --- | --- |
information,acceptingthe defaultvalueoffered,or cancelingthe application.
Ifnoactionis taken,theprogressof theapplicationishalted.Aresponsemay
eitherbe textual oraselection.The classificationfor Prompt messagesis%CP
• Info—ThisVB API messageispreceded bya bluedot.Infomessagetypes
containinformationsuch asuserrequests orfeedback from VBAPI orCreo
| Parametric.The |     | classificationfor | Infomessagesis |     | %CI |
| -------------- | --- | ----------------- | -------------- | --- | --- |
Note
Do notclassifymessagesthatdisplayinformationregarding problemswith
|     | an operationorprocess |           | asInfo. | Thesetypesofmessagesmustbe |     |
| --- | --------------------- | --------- | ------- | -------------------------- | --- |
|     | classifiedas          | Warnings. |         |                            |     |
• Warning—ThisVB APImessageis precededby atrianglecontainingan
exclamationpoint.Warningmessagetypescontaininformation toalert users
tosituations thatcouldpotentiallyleadto an errorduring alaterstageof the
process.Examplesofwarnings couldbe aprocessrestrictionor asuspected
dataproblem.AWarning willnot preventorinterruptaprocess.Also, a
Warning shouldnotbe usedto indicateafailedoperation.Warningsmustonly
cautionauserthatthe completedoperationmaynothavebeenperformedin a
completelydesirableway. Theclassification forWarningmessagesis %CW
• Error—ThisVB APImessageis precededby aa brokensquare.AnError
messageinforms theuserthatarequiredtask was notcompletedsuccessfully.
| Dependingonthe |     | application,afailedtaskmayor |     |     | maynot require |
| -------------- | --- | ---------------------------- | --- | --- | -------------- |
interventionor correctionbeforework cancontinue. Wheneverpossible
SessionObjects 51

| redressthissituationby |     |     | providinga |     | path.The | classificationfor | Error |
| ---------------------- | --- | --- | ---------- | --- | -------- | ----------------- | ----- |
messagesis%CE
| • Critical—This |     |     | VBAPI | messageispreceded |     | by aredX. | ACritical |
| --------------- | --- | --- | ----- | ----------------- | --- | --------- | --------- |
messagetypeinformsthe userofan extremelyserioussituationthatisusually
precededby lossofuserdata. Optionsredressing thissituation,if available,
shouldbe providedwithin themessage.Theclassification foraCritical
messagesis%CC
| Example | Code: | Writing |     | a Message |     |     |     |
| ------- | ----- | ------- | --- | --------- | --- | --- | --- |
Thesamplecodein thefilepfcSessionObjectsExamples.vb locatedat
| <creo_vbapi_loadpoint>/vbapi_examplesdemonstrates |                   |     |     |                   |     |             | howtowrite |
| ------------------------------------------------- | ----------------- | --- | --- | ----------------- | --- | ----------- | ---------- |
| amessageto                                        | themessagewindow. |     |     | Theprogramusesthe |     | messagefile |            |
mymessages.txt.
| Reading | Data | from | the | Message | Window |     |     |
| ------- | ---- | ---- | --- | ------- | ------ | --- | --- |
MethodsIntroduced:
• IpfcSession.UIReadIntMessage()
• IpfcSession.UIReadRealMessage()
• IpfcSession.UIReadStringMessage()
Thesemethodsenableaprogramto get datafrom theuser.The methodsobtain
| keyboardinputfrom |     | atextboxin |     | theCreo | Parametricuserinterface. |     |     |
| ----------------- | --- | ---------- | --- | ------- | ------------------------ | --- | --- |
Note
Whenthe userpressesEsc orclicks Cancelinthe CreoParametricuser
| interface,thesemethodsthrow |     |     |     | theexception |     |     |     |
| --------------------------- | --- | --- | --- | ------------ | --- | --- | --- |
IpfcXToolkitMsgUserQuit.
| TheIpfcSession.UIReadIntMessage()             |     |     |     |     |     | and |          |
| --------------------------------------------- | --- | --- | --- | --- | --- | --- | -------- |
| IpfcSession.UIReadRealMessage()methodscontain |     |     |     |     |     |     | optional |
argumentsthatcan beused tolimit thevalueof thedatato acertainrange.
| ThemethodIpfcSession.UIReadStringMessage() |     |     |     |     |     |     | includesan |
| ------------------------------------------ | --- | --- | --- | --- | --- | --- | ---------- |
optionalBoolean argument thatspecifieswhetherto echocharacters enteredonto
thescreen.You wouldusethisargumentwhenprompting auserto entera
password.
| 52  |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

Note
Adefaultvalue isdisplayedin thetextboxasinput.Whenuserpresses the
Enterkey asinputin theuserinterface,the defaultvalueisnot passedtothe
VB APImethod;instead,themethodreturns aconstantstringuse_
default_string. Whenthis stringisreturned,theapplicationmust
interpretthattheuserwants to usethedefaultvalue.
Displaying Feature Parameters
MethodIntroduced:
• pfcSession.Session.UIDisplayFeatureParams
ThemethodIpfcSession.UIDisplayFeatureParams() forcesCreo
Parametrictoshowdimensions orotherparametersstoredona specificfeature.
Thedisplayeddimensionsmaythenbe interactivelyselectedby theuser.
File Dialogs
MethodsandPropertiesIntroduced:
• IpfcSession.UIOpenFile()
• CCpfcFileOpenOptions.Create()
• IpfcFileOpenOptions.FilterString
• IpfcFileOpenOptions.PreselectedItem
• IpfcFileUIOptions.DefaultPath
• IpfcFileUIOptions.DialogLabel
• IpfcFileUIOptions.Shortcuts
• CCpfcFileOpenShortcut.Create()
• IpfcFileOpenShortcut.ShortcutName
• IpfcFileOpenShortcut.ShortcutPath
• IpfcSession.UISaveFile()
• CCpfcFileSaveOptions.Create()
• IpfcSession.UISelectDirectory()
• CCpfcDirectorySelectionOptions.Create()
• IpfcBaseSession.UIRegisterFileOpen
• CCpfcFileOpenRegisterOptions.Create
SessionObjects 53

• IpfcFileOpenRegisterOptions.FileDescription
• IpfcFileOpenRegisterOptions.FileType
• IpfcFileOpenRegisterListener.FileOpenAccess
• IpfcFileOpenRegisterListener.OnFileOpenRegister
• IpfcBaseSession.UIRegisterFileSave
• CCpfcFileSaveRegisterOptions.Create
• IpfcFileSaveRegisterOptions.FileDescription
• IpfcFileSaveRegisterOptions.FileType
• IpfcFileSaveRegisterListener.FileSaveAccess
• IpfcFileSaveRegisterListener.OnFileSaveRegister
ThemethodIpfcSession.UIOpenFile()
opens therelevantdialogboxfor
browsingdirectoriesandopeningfiles. Themethodlets youspecify several
| optionsthrough | theinput argumentsIpfcFileOpenOptions |     |     | and |     |
| -------------- | ------------------------------------- | --- | --- | --- | --- |
IpfcFileUIOptions.
| Usethe methodCCpfcFileOpenOptions.Create()to |     |     |     | createanew |     |
| -------------------------------------------- | --- | --- | --- | ---------- | --- |
instanceof theIpfcFileOpenOptions object.Thisobjectcontains the
followingoptions:
• FilterString—Specifiesthefilter stringforthe typeoffileaccepted by
| thedialogbox.                                      | Multiplefiletypesshouldbe |                                     | listed withwildcards | and |     |
| -------------------------------------------------- | ------------------------- | ----------------------------------- | -------------------- | --- | --- |
| separatedbycommas,for                              |                           | example,*.prt,*.asm,*.txt,*.avi,and |                      |     | so  |
| on.Use thepropertyIpfcFileOpenOptions.FilterString |                           |                                     |                      | to  | set |
thisoption.
• PreselectedItem—Specifies thenameof anitem to preselectinthe
| dialogbox.Use                       | theproperty |                   |                   |                   |     |
| ----------------------------------- | ----------- | ----------------- | ----------------- | ----------------- | --- |
| IpfcFileOpenOptions.PreselectedItem |             |                   |                   | tosetthis option. |     |
| TheIpfcFileUIOptions                |             | objectcontainsthe | followingoptions: |                   |     |
• DefaultPath—Specifiesthe nameofthe pathtobe opened bydefault in
| thedialogbox. | Usethe | propertyIpfcFileUIOptions.DefaultPathto |     |     |     |
| ------------- | ------ | --------------------------------------- | --- | --- | --- |
setthisoption.
• DialogLabel—Specifiesthe titleof thedialogbox. Usethe property
| IpfcFileUIOptions.DialogLabelto            |     |                           | setthisoption.     |           |     |
| ------------------------------------------ | --- | ------------------------- | ------------------ | --------- | --- |
| • Shortcuts—Specifiesan                    |     | arrayoffileshortcutsofthe |                    | type      |     |
| IpfcFileOpenShortcut.Createthisobjectusing |     |                           |                    | themethod |     |
| CCpfcFileOpenShortcut.Create().            |     |                           | Thisobjectcontains | the       |     |
followingattributes:
| ○ ShortcutName—Specifiesthename |            |     | ofshortcutpathto                | be made |     |
| ------------------------------- | ---------- | --- | ------------------------------- | ------- | --- |
| availableinthe                  | dialogbox. |     |                                 |         |     |
| 54                              |            |     | Creo®ParametricVBAPIUser’sGuide |         |     |

|     | ○ ShortcutPath—Specifiesthestring |                                       |     |     |     | fortheshortcut |     | path.         |     |     |
| --- | --------------------------------- | ------------------------------------- | --- | --- | --- | -------------- | --- | ------------- | --- | --- |
|     | Usethe                            | propertyIpfcFileUIOptions.Shortcutsto |     |     |     |                |     | setthearrayof |     |     |
fileshortcuts.
ThemethodIpfcSession.UIOpenFile() returnsthe fileselectedby you.
Theapplicationmustuseothermethods ortechniques toperformthe desired
| actionon |     | thefile. |     |     |     |     |     |     |     |     |
| -------- | --- | -------- | --- | --- | --- | --- | --- | --- | --- | --- |
ThemethodIpfcSession.UISaveFile() opens therelevantdialogboxfor
| savingafile.                    |     | Themethodaccepts |     | optionssimilar |                        | to                 |     |     |       |     |
| ------------------------------- | --- | ---------------- | --- | -------------- | ---------------------- | ------------------ | --- | --- | ----- | --- |
| IpfcSession.UIOpenFile()through |     |                  |     |                | theIpfcFileSaveOptions |                    |     |     |       | and |
| IpfcFileUIOptionsobjects.Use    |     |                  |     | themethod      |                        |                    |     |     |       |     |
| CCpfcFileSaveOptions.Create()to |     |                  |     |                |                        | createanewinstance |     |     | ofthe |     |
IpfcFileSaveOptions object.Whenusing the Savedialogbox, youcan set
thenameto anon-existentfile. Themethod IpfcSession.UISaveFile()
returnsthename ofthefileselectedbyyou; theapplicationmustuseother
| methodsortechniques                      |     |     | toperformthe | desiredactionon |     |     | thefile. |                  |     |     |
| ---------------------------------------- | --- | --- | ------------ | --------------- | --- | --- | -------- | ---------------- | --- | --- |
| ThemethodIpfcSession.UISelectDirectory() |     |     |              |                 |     |     |          | promptstheuserto |     |     |
selecta directoryusing theCreo Parametricdialogboxfor browsingdirectories.
| Themethodacceptsoptions                           |     |     | through | the |     |           |     |     |        |     |
| ------------------------------------------------- | --- | --- | ------- | --- | --- | --------- | --- | --- | ------ | --- |
| IpfcDirectorySelectionOptionsobjectwhichissimilar |     |     |         |     |     |           |     |     | to the |     |
| IpfcFileUIOptionsobject(describedfor              |     |     |         |     |     | themethod |     |     |        |     |
IpfcSession.UIOpenFile()).Specifythedefault
directorypath,thetitle
ofthedialogbox, andasetof shortcutsto otherdirectoriesto startbrowsing.If
thedefaultpathisspecifiedasNULL, thecurrentdirectoryisused.Use the
| methodCCpfcDirectorySelectionOptions.Create() |     |                                            |     |     |     |     |     |     | tocreatea | new |
| --------------------------------------------- | --- | ------------------------------------------ | --- | --- | --- | --- | --- | --- | --------- | --- |
| instanceof                                    |     | theIpfcDirectorySelectionOptionsobject.The |     |     |     |     |     |     | method    |     |
IpfcSession.UISelectDirectory()returnsthe selecteddirectorypath;
theapplicationmustuseothermethodsor techniquesto performotherrelevant
| taskswith                                     |     | thisselectedpath. |     |     |     |     |     |     |               |     |
| --------------------------------------------- | --- | ----------------- | --- | --- | --- | --- | --- | --- | ------------- | --- |
| ThemethodIpfcBaseSession.UIRegisterFileOpen() |     |                   |     |     |     |     |     |     | registersanew |     |
filetypein theFile▶▶Opendialogboxin Creo Parametric.This methodtakesthe
IpfcFileOpenRegisterOptionsand
IpfcFileOpenRegisterListenerobjects asitsinputarguments. These
objectsareasfollows:
• IpfcFileOpenRegisterOptions—This objectcontainsthe options for
|     | registeringan                       | open | operation.Use                | themethod |     |     |              |     |          |     |
| --- | ----------------------------------- | ---- | ---------------------------- | --------- | --- | --- | ------------ | --- | -------- | --- |
|     | CCpfcFileOpenRegisterOptions.Create |      |                              |           |     |     | tocreateanew |     | instance |     |
|     | oftheobject.It                      |      | containsthefollowingoptions: |           |     |     |              |     |          |     |
○ FileDescription—Specifies theshort descriptionof thefiletypeto
|     | be                                          | opened.This | descriptionappears |     | for | thefiletypeinthe |     | File▶▶Open |     |        |
| --- | ------------------------------------------- | ----------- | ------------------ | --- | --- | ---------------- | --- | ---------- | --- | ------ |
|     | dialogbox.                                  |             | Usethe property    |     |     |                  |     |            |     |        |
|     | IpfcFileOpenRegisterOptions.FileDescription |             |                    |     |     |                  |     |            | to  | modify |
thisoption.
| SessionObjects |     |     |     |     |     |     |     |     |     | 55  |
| -------------- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |

○ FileType—Specifiesthefiletypeto beopened. Thefiletypeappearsas
|     | thefileextensionin                   |     |     | theFile▶▶Opendialogbox. |     |     | Use | theproperty   |     |
| --- | ------------------------------------ | --- | --- | ----------------------- | --- | --- | --- | ------------- | --- |
|     | IpfcFileOpenRegisterOptions.FileType |     |     |                         |     |     |     | to modifythis |     |
option.
| •   | IpfcFileOpenRegisterListener—This              |     |     |                  |     |                  | objectprovidesthe |        | action   |
| --- | ---------------------------------------------- | --- | --- | ---------------- | --- | ---------------- | ----------------- | ------ | -------- |
|     | listenermethodsfor                             |     |     | thenewfiletypeto |     | beregistered.The |                   | method |          |
|     | Ipfc.FileOpenRegisterListener.FileOpenAccessis |     |     |                  |     |                  |                   |        | calledto |
determinewhetherthe newfiletypecan beopened using theFile▶▶Open
|                                               | dialogbox.The                                   |     | method |                             |     |     |     |     |               |
| --------------------------------------------- | ----------------------------------------------- | --- | ------ | --------------------------- | --- | --- | --- | --- | ------------- |
|                                               | IpfcFileOpenRegisterListener.OnFileOpenRegister |     |        |                             |     |     |     |     | is            |
|                                               | calledonclickingOpenfor                         |     |        | thenewlyregisteredfiletype. |     |     |     |     |               |
| ThemethodIpfcBaseSession.UIRegisterFileSave() |                                                 |     |        |                             |     |     |     |     | registersanew |
filetypein theFile▶▶SaveaCopydialogboxin Creo Parametric.Thismethod
takestheIpfcFileSaveRegisterOptionsand
IpfcFileSaveRegisterListenerobjects asitsinputarguments. These
| objectsaredescribedas |     |     |     | follows: |     |     |     |     |     |
| --------------------- | --- | --- | --- | -------- | --- | --- | --- | --- | --- |
• IpfcFileSaveRegisterOptions—This objectcontainsthe options for
|     | registeringa                          | saveoperation.Usethe |     |     |                      | method |     |           |     |
| --- | ------------------------------------- | -------------------- | --- | --- | -------------------- | ------ | --- | --------- | --- |
|     | CCpfcFileSaveRegisterOptions_Create() |                      |     |     |                      |        |     | tocreatea | new |
|     | instanceof                            | theobject.Itcontains |     |     | thefollowingoptions: |        |     |           |     |
○ FileDescription—Specifies theshort descriptionof thefiletypeto
|     | be saved.This                               |     | descriptionappearsfor |        |          | thefiletypeinthe |     | File▶▶Savea |           |
| --- | ------------------------------------------- | --- | --------------------- | ------ | -------- | ---------------- | --- | ----------- | --------- |
|     | Copydialogbox.                              |     |                       | Usethe | property |                  |     |             |           |
|     | IpfcFileSaveRegisterOptions.FileDescription |     |                       |        |          |                  |     |             | to modify |
thisoption.
○ FileType—Specifiesthefiletypeto besaved. Thefiletypeappears as
thefileextensionin theFile▶▶SaveaCopydialogbox. Use theproperty
|     | IpfcFileSaveRegisterOptions.FileType |     |     |     |     |     |     | to modifythis |     |
| --- | ------------------------------------ | --- | --- | --- | --- | --- | --- | ------------- | --- |
option.
| •   | IpfcFileSaveRegisterListener—This             |     |     |                  |     |                  | objectprovidesthe |        | action     |
| --- | --------------------------------------------- | --- | --- | ---------------- | --- | ---------------- | ----------------- | ------ | ---------- |
|     | listenermethodsfor                            |     |     | thenewfiletypeto |     | beregistered.The |                   | method |            |
|     | IpfcFileSaveRegisterListener.FileSaveAccess() |     |     |                  |     |                  |                   |        | iscalledto |
determinewhetherthe newfiletypecan besavedusing theFile▶▶SaveaCopy
|     | dialogbox.The                                   |     | method |                                 |     |     |                                 |     |     |
| --- | ----------------------------------------------- | --- | ------ | ------------------------------- | --- | --- | ------------------------------- | --- | --- |
|     | IpfcFileSaveRegisterListener.OnFileSaveRegister |     |        |                                 |     |     |                                 |     | is  |
|     | calledonclickingOK                              |     |        | forthe newlyregisteredfiletype. |     |     |                                 |     |     |
| 56  |                                                 |     |        |                                 |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |

| Customizing |     |     | the Creo | Parametric |     | Navigation |     | Area |
| ----------- | --- | --- | -------- | ---------- | --- | ---------- | --- | ---- |
TheCreo Parametricnavigationareaincludes theModeland LayerTreepane,
Folderbrowserpane,andFavorites pane. Themethodsdescribedin thissection
enablethe VBAPI applicationstoadd custompanesthatcontain Web pagesto the
CreoParametricnavigationarea.
| Adding | Custom |     | Web Pages |     |     |     |     |     |
| ------ | ------ | --- | --------- | --- | --- | --- | --- | --- |
ToaddcustomWeb pagestothe navigationarea,thethe VBAPI application
must:
| 1. Addanewpaneto |             |       | thenavigationarea. |     |             |          |     |     |
| ---------------- | ----------- | ----- | ------------------ | --- | ----------- | -------- | --- | --- |
| 2. Setan         | iconforthis |       | pane.              |     |             |          |     |     |
| 3. SettheURL     |             | ofthe | locationthatwillbe |     | displayedin | thepane. |     |     |
MethodsIntroduced:
• IpfcSession.NavigatorPaneBrowserAdd()
• IpfcSession.NavigatorPaneBrowserIconSet()
• IpfcSession.NavigatorPaneBrowserURLSet()
| ThemethodIpfcSession.NavigatorPaneBrowserAdd() |     |     |     |     |     |     |     | addsa new |
| ---------------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --------- |
panethatcan displaya Web pageto thenavigationarea.The inputparametersare:
| • PaneName—Specifyaunique |     |     |     | nameforthe | pane.Use | thisname |     | in  |
| ------------------------- | --- | --- | --- | ---------- | -------- | -------- | --- | --- |
susbsequentcallsto
| IpfcSession.NavigatorPaneBrowserIconSet() |     |     |     |     |     |     | and |     |
| ----------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- |
IpfcSession.NavigatorPaneBrowserURLSet().
• IconFileName—Specifythename oftheiconfile,including theextension. A
validformatfor theiconfileis thePTC-proprietary formatused byCreo
Parametric.BIF,.GIF, .JPG, or.PNG. Thenewpaneisdisplayedwith the
iconimage.Ifyouspecify thevalue asNULL, thedefaultCreo Parametricicon
isused.
| Thedefaultsearch |                                    |     | pathsforfindingtheiconsare: |     |     |             |     |     |
| ---------------- | ---------------------------------- | --- | --------------------------- | --- | --- | ----------- | --- | --- |
| ○                | <creo_loadpoint>\<datecode>\Common |     |                             |     |     | Files\text\ |     |     |
resource
○ <Applicationtextdir>\resource
○ <Applicationtextdir>\<language>\resource
Thelocationof theapplicationtextdirectoryis specifiedinthe registryfile.
| • URL—SpecifytheURL |                                                 |     | ofthelocationto |     | be  | accessedfrom | thepane. |     |
| ------------------- | ----------------------------------------------- | --- | --------------- | --- | --- | ------------ | -------- | --- |
| Usethe              | methodIpfcSession.NavigatorPaneBrowserIconSet() |     |                 |     |     |              |          | to  |
setor changetheiconofaspecifiedbrowserpanein thenavigationarea.
SessionObjects 57

Usethe methodIpfcSession.NavigatorPaneBrowserURLSet() to
changethe URLof thepagedisplayedinthe browserpanein thenavigationarea.
| Example: | Customizing | the Creo | Parametric | Navigation |
| -------- | ----------- | -------- | ---------- | ---------- |
Pane
| Thesamplecodein | thefilepfcNavigatorPane.vb |     |     | locatedat<creo_ |
| --------------- | -------------------------- | --- | --- | --------------- |
vbapi_loadpoint>/vbapi_examplesshowsyouhowto customizethe
CreoParametricnavigationpane.
| 58  |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

6
Selection
InteractiveSelection..................................................................................................60
AccessingSelectionData..........................................................................................61
ProgrammaticSelection.............................................................................................63
SelectionBuffer.........................................................................................................64
Thischapterdescribes howtouseInteractiveSelection inthe VBAPI.
59

| Interactive Selection |     |     |     |     |
| --------------------- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcBaseSession.Select()
• CCpfcSelectionOptions.Create()
• IpfcSelectionOptions.MaxNumSels
• IpfcSelectionOptions.OptionKeywords
| ThemethodIpfcBaseSession.Select() |     | activatesthe | standardCreo |     |
| --------------------------------- | --- | ------------ | ------------ | --- |
Parametricmenustructurefor selectingobjectsandreturns aIpfcSelection
sequencethatcontainsthe objectstheuserselected. Usingthe Optionsargument,
youcancontrolthetypeofobjectthatcan beselectedandthe maximumnumber
ofselections.
Inaddition,youcan passina IpfcSelection sequenceto themethod.The
returnedIpfcSelection sequencewillcontainthe inputsequenceandany
newobjects.
| ThemethodCCpfcSelectionOptions.Create()  |     |     | and theproperty |      |
| ---------------------------------------- | --- | --- | --------------- | ---- |
| IpfcSelectionOptions.OptionKeywordstakea |     |     | Stringargument  | made |
upofone ormoreofthe identifierslisted in thetable below,separatedby commas.
Forexample,to allowtheselectionoffeaturesandaxes,theargumentswouldbe
feature,axis.
| CreoParametricDatabase | StringIdentifier | ModelItemType |     |     |
| ---------------------- | ---------------- | ------------- | --- | --- |
Item
| Datumpoint | point | EpfcITEM_POINT |     |     |
| ---------- | ----- | -------------- | --- | --- |
| Datumaxis  | axis  | EpfcITEM_AXIS  |     |     |
EpfcITEM_SURFACE
| Datumplane                | datum   |                    |     |     |
| ------------------------- | ------- | ------------------ | --- | --- |
| Coordinatesystemdatum     | csys    | EpfcITEM_COORD_SYS |     |     |
| Feature                   | feature | EpfcITEM_FEATURE   |     |     |
| Edge(solidordatumsurface) | edge    | EpfcITEM_EDGE      |     |     |
| Edge(solidonly)           | sldedge | EpfcITEM_EDGE      |     |     |
| Edge(datumsurfaceonly)    | qltedge | EpfcITEM_EDGE      |     |     |
|                           | curve   | EpfcITEM_CURVE     |     |     |
Datumcurve
| Compositecurve        | comp_crv | EpfcITEM_CURVE   |     |     |
| --------------------- | -------- | ---------------- | --- | --- |
| Surface(solidorquilt) | surface  | EpfcITEM_SURFACE |     |     |
| Surface(solid)        | sldface  | EpfcITEM_SURFACE |     |     |
| Surface(datumsurface) | qltface  | EpfcITEM_SURFACE |     |     |
ITEM_QUILT
| Quilt              | dtmqlt    |                        |     |     |
| ------------------ | --------- | ---------------------- | --- | --- |
| Dimension          | dimension | EpfcITEM_DIMENSION     |     |     |
| Referencedimension | ref_dim   | EpfcITEM_REF_DIMENSION |     |     |
| Integerparameter   | ipar      | EpfcITEM_DIMENSION     |     |     |
| Part               | part      | N/A                    |     |     |
prt_or_asm
| Partorsubassembly |     | N/A                             |     |     |
| ----------------- | --- | ------------------------------- | --- | --- |
| 60                |     | Creo®ParametricVBAPIUser’sGuide |     |     |

| CreoParametricDatabase | StringIdentifier |     | ModelItemType |     |
| ---------------------- | ---------------- | --- | ------------- | --- |
Item
component
| Assemblycomponentmodel |            |     | N/A               |     |
| ---------------------- | ---------- | --- | ----------------- | --- |
| Componentorfeature     | membfeat   |     | EpfcITEM_FEATURE  |     |
| Detailsymbol           | dtl_symbol |     | EpfcITEM_DTL_SYM_ |     |
INSTANCE
|     | any_note |     | EpfcITEM_NOTE,EpfcITEM_ |     |
| --- | -------- | --- | ----------------------- | --- |
Note
DTL_NOTE
EpfcITEM_DTL_ENTITY
| Draftentity   | draft_ent  |     |                      |     |
| ------------- | ---------- | --- | -------------------- | --- |
| Table         | dwg_table  |     | EpfcITEM_TABLE       |     |
| Tablecell     | table_cell |     | EpfcITEM_TABLE       |     |
| Drawingview   | dwg_view   |     | N/A                  |     |
| Solidbody     | 3d_body    |     | EpfcITEM_BODY        |     |
| DatumCurveEnd | curve_end  |     | EpfcITEM_CRV_STARTor |     |
EpfcITEM_CRV_END
| Whenyouspecifythe                     | maximumnumberofselections,theargumentto |     |            |     |
| ------------------------------------- | --------------------------------------- | --- | ---------- | --- |
| IpfcSelectionOptions.MaxNumSelsmustbe |                                         |     | anInteger. |     |
Thedefaultvalue assignedwhencreatingaIpfcSelectionOptions object
| is–1,whichallows | anynumberofselectionsby | theuser. |     |     |
| ---------------- | ----------------------- | -------- | --- | --- |
| Accessing        | Selection               | Data     |     |     |
PropertiesIntroduced:
• IpfcSelection.SelModel
• IpfcSelection.SelItem
• IpfcSelection.Path
• IpfcSelection.Params
• IpfcSelection.TParam
• IpfcSelection.Point
• IpfcSelection.Depth
• IpfcSelection.SelView2D
• IpfcSelection.SelTableCell
• IpfcSelection.SelTableSegment
Thesepropertiesreturnobjectsanddatathatmakeup theselectionobject.Using
| theappropriateproperties,youcan | accessthefollowingdata:            |     |     |     |
| ------------------------------- | ---------------------------------- | --- | --- | --- |
| • Fora selectedmodelor          | modelitemuseIpfcSelection.SelModel |     |     | or  |
IpfcSelection.SelItem.
| • Foran assembly | componentuseIpfcSelection.Path. |     |     |     |
| ---------------- | ------------------------------- | --- | --- | --- |
| Selection        |                                 |     |     | 61  |

| • ForUV | parametersof |     | theselectionpoint |     |     | onasurface | use |
| ------- | ------------ | --- | ----------------- | --- | --- | ---------- | --- |
IpfcSelection.Params.
| • Forthe | Tparameter |     | ofthe | selectionpointon |     | anedgeor | curveuse |
| -------- | ---------- | --- | ----- | ---------------- | --- | -------- | -------- |
IpfcSelection.TParam.
• Fora three-dimensionalpointobjectthatcontains theselectedpointuse
IpfcSelection.Point.
• Forselection depth,inscreen coordinatesuseIpfcSelection.Depth.
• Forthe selecteddrawingview,if theselection wasfrom adrawing, use
IpfcSelection.SelView2D.
| • Forthe | selectedtablecell,ifthe |     |     |     | selectionwas | froma | table,use |
| -------- | ----------------------- | --- | --- | --- | ------------ | ----- | --------- |
IpfcSelection.SelTableCell.
• Forthe selectedtablesegment,if theselectionwas from atable,use
IpfcSelection.SelTableSegment.
| Controlling |     | Selection |     |     | Display |     |     |
| ----------- | --- | --------- | --- | --- | ------- | --- | --- |
MethodsIntroduced:
• IpfcSelection.Highlight()
• IpfcSelection.UnHighlight()
• IpfcSelection.Display()
Thesemethodscauseaspecific selectionto behighlighted ordimmedon the
| screenusingthe |     | colorspecifiedasan |     |     | argument. |     |     |
| -------------- | --- | ------------------ | --- | --- | --------- | --- | --- |
ThemethodIpfcSelection.Highlight() highlightstheselection inthe
currentwindow.This highlightisthe sameastheone usedbyCreo Parametric
whenselectingan item—it justrepaintsthe wire-framedisplayinthe newcolor.
| Thehighlightis                      |     | removedif |     | youusetheView,Repaintcommandor |     |          |     |
| ----------------------------------- | --- | --------- | --- | ------------------------------ | --- | -------- | --- |
| IpfcWindow.Repaint();itisnotremoved |     |           |     |                                |     | ifyouuse |     |
IpfcWindow.Refresh().
| ThemethodIpfcSelection.UnHighlight() |                  |     |     |                         |     | removesthehighlight.      |     |
| ------------------------------------ | ---------------- | --- | --- | ----------------------- | --- | ------------------------- | --- |
| ThemethodIpfcSelection.Display()     |                  |     |     |                         |     | causesaselectedobjecttobe |     |
| displayedon                          | thescreen,evenif |     |     | itissuppressedorhidden. |     |                           |     |
Note
| This | isaone-time |     | actionand | thenextrepaintwillerasethisdisplay. |     |                                 |     |
| ---- | ----------- | --- | --------- | ----------------------------------- | --- | ------------------------------- | --- |
| 62   |             |     |           |                                     |     | Creo®ParametricVBAPIUser’sGuide |     |

| Example | Code: | Using | Interactive | Selection |     |
| ------- | ----- | ----- | ----------- | --------- | --- |
Thesamplecodein thefilepfcSelectionExamples.vblocatedat<creo_
vbapi_loadpoint>/vbapi_examplesdemonstrateshowto
usetheVB
APIto allowinteractiveselection.
| Programmatic |                           | Selection |        |                      |     |
| ------------ | ------------------------- | --------- | ------ | -------------------- | --- |
| TheVB        | APIprovidesmethodswhereby |           | youcan | makeyourownSelection |     |
objects,withoutprompting theuser. TheseSelections arerequiredasinputs to
somemethods andcanalso beused tohighlight certainobjectson thescreen.
MethodsandPropertiesIntroduced:
• CMpfcSelect.CreateModelItemSelection()
• CMpfcSelect.CreateComponentSelection()
• CMpfcSelect.CreateModelSelection()
• IpfcSelection.SelItem
• IpfcSelection.SelTableCell
• IpfcSelection.SetSelView2D
| ThemethodpCMpfcSelect.CreateModelItemSelection() |     |     |                               |     | createsa |
| ------------------------------------------------ | --- | --- | ----------------------------- | --- | -------- |
| selectionoutofanymodelitem                       |     |     | object.It takesaIpfcModelItem |     | and      |
optionallya IpfcComponentPath objecttoidentify whichcomponentin an
| assemblytheSelection                            |     | Objectbelongsto. |                      |          |          |
| ----------------------------------------------- | --- | ---------------- | -------------------- | -------- | -------- |
| ThemethodCMpfcSelect.CreateComponentSelection() |     |                  |                      |          | createsa |
| selectionoutofanycomponentin                    |     |                  | an assembly.         | Ittakesa |          |
| IpfcComponentPathobject.For                     |     |                  | moreinformationabout |          |          |
IpfcComponentPathobjects,seethe sectionGettingaSolidObjecton page
| 155inthe                     | Solidonpage154chapter.                   |     |                  |     |            |
| ---------------------------- | ---------------------------------------- | --- | ---------------- | --- | ---------- |
| Usethe                       | methodCMpfcSelect.CreateModelSelection() |     |                  |     | to createa |
| IpfcSelectionobject,basedona |                                          |     | IpfcModelobject. |     |            |
SomeVBAPI propertiesrequiremoreinformation to besetin theselection
| object.The                | methodsallowyouto |           | setthefollowing:          |     |     |
| ------------------------- | ----------------- | --------- | ------------------------- | --- | --- |
| Theselecteditemusing      |                   | themethod | IpfcSelection.SetSelItem. |     |     |
| Theselectedtablecellusing |                   | themethod |                           |     |     |
IpfcSelection.SetSelTableCell.
| Theselecteddrawing |     | viewusing | themethod |     |     |
| ------------------ | --- | --------- | --------- | --- | --- |
IpfcSelection.SetSelView2D.
Selection 63

| Selection    |     | Buffer |           |     |         |     |
| ------------ | --- | ------ | --------- | --- | ------- | --- |
| Introduction |     | to     | Selection |     | Buffers |     |
Selectionisthe processofchoosingitems on whichyouwantto performan
operation.In CreoParametric,beforea featuretoolis invoked,theusercanselect
itemsto be usedin agiventool'scollectors.Collectors arelikestoragebins ofthe
referencesof selecteditems.Thelocationwherepreselecteditems arestoredis
calledtheselectionbuffer.
Dependingon thesituation,different selectionbuffersmaybe activeatanyone
time.InPart andAssemblymode,CreoParametric offersthedefaultselection
buffer,the Editselectionbuffer, andothermorespecializedbuffers. OtherCreo
| Parametricmodesoffer |     |     | differentselection |     | buffers. |     |
| -------------------- | --- | --- | ------------------ | --- | -------- | --- |
InthedefaultPart andAssemblybuffer therearetwolevels atwhichselectionis
done:
• FirstLevelSelection
Provides accesstohigher-level objectssuchasfeaturesor components.You
canmakea secondlevel selectiononlyafteryouselectthehigher-level object.
• SecondLevelSelection
| Provides | accesstogeometricobjectssuch |     |     |     | asedgesandfaces. |     |
| -------- | ---------------------------- | --- | --- | --- | ---------------- | --- |
Note
First-level andsecond-levelobjectsareusuallyincompatibleinthe selection
buffer.
TheVB APIallowsaccess tothe contentsof thecurrentlyactiveselectionbuffer.
| Theavailablefunctionsallow |            |               | yourapplicationto:        |                        |               |        |
| -------------------------- | ---------- | ------------- | ------------------------- | ---------------------- | ------------- | ------ |
| • Getthecontents           |            | ofthe         | activeselectionbuffer.    |                        |               |        |
| • Removethe                |            | contentsofthe |                           | activeselectionbuffer. |               |        |
| • Addtothe                 | contentsof |               | theactiveselectionbuffer. |                        |               |        |
| Reading                    | the        | Contents      |                           | of                     | the Selection | Buffer |
PropertiesIntroduced:
• IpfcSession.CurrentSelectionBuffer
• IpfcSelectionBuffer.Contents
| 64  |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

| ThepropertyIpfcSession.CurrentSelectionBuffer |     |     |     | returns the |
| --------------------------------------------- | --- | --- | --- | ----------- |
selectionbufferobjectfor thecurrentactivemodelin session.The selectionbuffer
containstheitems preselectedby theuserto beused bythe selectiontooland
popupmenus.
Usethe propertyIpfcSelectionBuffer.Contents toaccess thecontents
ofthecurrentselectionbuffer. Themethodreturns independentcopies ofthe
selectionsin theselectionbuffer (if thebuffer iscleared,thisarrayis stillvalid).
| Removing | the | Items of | the Selection | Buffer |
| -------- | --- | -------- | ------------- | ------ |
MethodsIntroduced:
• IpfcSelectionBuffer.RemoveSelection()
• IpfcSelectionBuffer.Clear()
methodIpfcSelectionBuffer.RemoveSelection()
Usethe toremove
aspecificselectionfrom theselection buffer.The inputargument isthe
IndexToRemovespecifiesthe indexwheretheitem wasfound in thecallto the
propertyIpfcSelectionBuffer.Contents.
Usethe methodIpfcSelectionBuffer.Clear()to clearthe currently
activeselectionbuffer ofallcontents. Afterthe buffer is cleared,allcontents are
lost.
| Adding | Items | to the Selection | Buffer |     |
| ------ | ----- | ---------------- | ------ | --- |
MethodIntroduced:
• IpfcSelectionBuffer.AddSelection()
Usethe methodIpfcSelectionBuffer.AddSelection() to addan item
| tothe currentlyactive |     | selectionbuffer. |     |     |
| --------------------- | --- | ---------------- | --- | --- |
Note
The selecteditem mustreferto an itemthatis inthe currentmodelsuch asits
| owner,                 | componentpathor                 | drawingview. |                   |     |
| ---------------------- | ------------------------------- | ------------ | ----------------- | --- |
| Thismethodmayfaildueto |                                 | anyofthe     | followingreasons: |     |
| • Thereis              | nocurrentselectionbufferactive. |              |                   |     |
| • Theselection         | doesnotrefer                    | tothe        | currentmodel.     |     |
• Theitem isnotcurrentlydisplayedandso cannotbe addedto thebuffer.
Selection 65

• Theselection cannotbe addedto thebuffer in combinationwithone ormore
objectsthatarealreadyin thebuffer. Forexample: geometryandfeatures
| cannotbeselectedin | thedefaultbuffer | atthesametime. |
| ------------------ | ---------------- | -------------- |
66 Creo®ParametricVBAPIUser’sGuide

7
Ribbon Tabs, Groups, and Menu
Items
CreatingRibbonTabs,Groups,andMenuItems..........................................................68
AbouttheRibbonDefinitionFile..................................................................................70
LocalizingtheRibbonUserInterfaceCreatedbytheVBAPIApplications......................73
SupportforLegacytheVBAPIApplications................................................................74
MigrationofLegacyVBAPIApplications.....................................................................75
Thischapterdescribes theVB APIsupportfor theRibbon UserInterface(UI).It
alsodescribesthe impactof theribbonuserinterfaceon legacyVB API
applicationsandthe procedureto place thecommands,buttons, andmenuitems
createdbythe legacyapplicationsin the CreoParametricribbonuser interface.
Refertothe CreoParametricHelp formoreinformation onthe ribbonuser
interfaceandtheprocedureto customizethe ribbon.
67

| Creating | Ribbon |     | Tabs, | Groups, |     | and | Menu |
| -------- | ------ | --- | ----- | ------- | --- | --- | ---- |
Items
Customizationsto theribbon userinterfaceusing theVB APIapplicationsare
supportedthrough theCustomizeRibbontab inthe CreoParametricCreo
prametricOptions dialogbox. Youcan specifythe userinterfacelayoutfora VB
APIapplicationandsavethe layoutdefinitionin aribbondefinitionfile,
| toolkitribbonui.rbn.Setthe |     |     |     | configurationoptiontk_enable_ |     |     |     |
| -------------------------- | --- | --- | --- | ----------------------------- | --- | --- | --- |
ribbon_custom_saveto yesbeforecustomizing theribbon userinterface
| usingthe theVB | APIapplication.When |     |     | yourunCreo | Parametric,the |     |     |
| -------------- | ------------------- | --- | --- | ---------- | -------------- | --- | --- |
toolkitribbonui.rbn fileis loadedalong withthe VBAPI applicationand
thecommandscreated bytheVB API applicationappear in theribbonuser
interface.Referto thesection Aboutthe RibbonDefinition Fileonpage70 for
| moreinformationonthe |     | toolkitribbonui.rbn |     |     | file. |     |     |
| -------------------- | --- | ------------------- | --- | --- | ----- | --- | --- |
You cancustomizetheribbonuserinterfaceonly for aparticular modein Creo
Parametric.For example,ifyoucustomizetheribbonuser interfaceandsaveitto
| thetoolkitribbonui.rbn |     |     | fileinthe | Partmode,thenonloading |     |     | Creo |
| ---------------------- | --- | --- | --------- | ---------------------- | --- | --- | ---- |
Parametricthecustomizeduserinterfacewillbe visibleonlyin thePart mode.To
viewaparticular taborgroupin allthe modes,youmustcustomizethe ribbon
userinterfaceandsavethe toolkitribbonui.rbn filein eachmode.Refer
tothe CreoParametricFundamentalsHelp formoreinformation oncustomizing
theribbon.
Note
Youcanadd anewgroup toan existing taborcreatea newtabusing the
CustomizeRibbontab inthe CreoParametricOptions dialogbox.Youwillnot
be abletomodifythe tabsorgroups thataredefinedby CreoParametric.
| Workflow | to Add | Menu | Items | to  | the Ribbon |     | User |
| -------- | ------ | ---- | ----- | --- | ---------- | --- | ---- |
Interface
Note
The instructionsexplained belowareapplicableonly ifthe applicationis
implementedinfull asynchronousmode.This isbecauseapplicationsin
| simpleasynchronousmode |     |     | cannothandlerequests, |     | thatis,command |     |     |
| ---------------------- | --- | --- | --------------------- | --- | -------------- | --- | --- |
callbacks,from CreoParametric.Referto thechapteron AsynchronousMode,
for moreinformation.
| 68  |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- | --- |

Settheconfigurationoptiontk_enable_ribbon_custom_save to yes
beforecustomizingthe ribbonuserinterface.Thesteps to addcommandsto the
CreoParametricribbonuserinterface areasfollows:
1. CreateaVB APIapplicationwithcomplete commanddefinition, which
includesspecifyingcommand label,helptext,largeiconname,andsmall icon
name.Designatethecommandusing IpfcUICommand.Designate().
2. StarttheVB APIapplicationandensurethatithasstartedorconnectedto
CreoParametric.The commandscreatedby theVB APIapplicationwillbe
loadedinCreo Parametric.
3. ClickFile▶▶Options.The CreoParametricOptions dialogboxopens.
4. ClickCustomizeRibbon.
5. Inthe CustomizetheRibbonlist, selecta tabandcreateanewgroup in itor
createanewtaband agroup init.
6. Inthe Choosecommandsfromlist,selectTOOLKITCommands.The
commandscreatedby theVB API applicationaredisplayed.
7. ClickAdd toadd thecommandsto thenewtab orgroup.
8. ClickImport/Export▶▶SavetheAuxilliaryApplicationUserInterface.The
changesaresavedtothe toolkitribbonui.rbn file.The
toolkitribbonui.rbn fileissavedinthe textfolderspecifiedin theVB
APIapplicationregistryfile.For moreinformationrefer tothe sectionon
AbouttheRibbon DefinitionFileon page70.
Note
The SavetheAuxilliaryApplicationUserInterfacebutton isenabled onlyif
yousetthe configurationoption tk_enable_ribbon_custom_save
to yes.
9. ClickApply.The customsettings aresavedto the toolkitribbonui.rbn
file.
10. ReloadtheVB APIapplicationorrestart CreoParametric.The
toolkitribbonui.rbn filewillbe loadedalongwith theVB API
application.
Iftranslatedmessagesareavailablefor thenewlyaddedtabs orgroups,thenCreo
Parametricdisplaysthe translatedstringsby searching forthe samestringfromthe
listofstringbasedmessages thatare loaded.For moreinformationrefer tothe
sectiononLocalizing theRibbon UserInterfaceCreated bythe VBAPI
Applicationsonpage73.
RibbonTabs,Groups,andMenuItems 69

| About |     | the | Ribbon |     |     | Definition |     | File |     |     |
| ----- | --- | --- | ------ | --- | --- | ---------- | --- | ---- | --- | --- |
Aribbondefinitionfileis afilethatis createdthrough theCustomizeRibbon
interfaceinCreo Parametric.Thisfiledefines thecontainers,thatis,Tabs,Group,
orCascademenusthatarecreated bya particularVB APIapplication. Itcontains
informationonwhetherto showan iconorlabel.It alsocontains thesizeof the
| iconto | be  | used,thatis,alarge |     |     | icon(32X32)ora |     | smallicon(16x16). |     |     |     |
| ------ | --- | ------------------ | --- | --- | -------------- | --- | ----------------- | --- | --- | --- |
Theribbonuserinterface displaysthe commandsreferencedin the ribbon
definitionfileonlyif thecommands areloadedandarevisiblein thatparticular
CreoParametricmode.Iftranslated messagesareavailableforthe newlyadded
tabsorgroups, thenCreo Parametricdisplaysthetranslated stringsbysearching
forthesamestring fromthelist ofstringbasedmessagesthatare loaded.For more
informationreferto thesection onLocalizing theRibbon UserInterfaceCreated
| bytheVB                                               |     | API | Applicationson |     | page73. |     |     |     |        |     |
| ----------------------------------------------------- | --- | --- | -------------- | --- | ------- | --- | --- | --- | ------ | --- |
| Settheconfigurationoptiontk_enable_ribbon_custom_save |     |     |                |     |         |     |     |     | to yes |     |
beforecustomizingthe ribbonuserinterface.To savetheribbon userinterface
| layoutdefinitionto |                        |     | thetoolkitribbonui.rbn |     |     |     |     | file:           |     |     |
| ------------------ | ---------------------- | --- | ---------------------- | --- | --- | --- | --- | --------------- | --- | --- |
| 1.                 | ClickFile▶▶Options.The |     |                        |     |     |     |     | dialogboxopens. |     |     |
CreoParametricOptions
2. ClickCustomizeRibbon.
3. Inthe CustomizetheRibbonlist, selecta tabandcreateanewgroup in itor
|     | createanewtaband  |                                                  |               | agroup | init. |                          |          |     |     |     |
| --- | ----------------- | ------------------------------------------------ | ------------- | ------ | ----- | ------------------------ | -------- | --- | --- | --- |
| 4.  | Inthe             | Choosecommandsfromlist,selectTOOLKITCommands.The |               |        |       |                          |          |     |     |     |
|     | commandscreatedby |                                                  |               | theVB  | API   | applicationaredisplayed. |          |     |     |     |
| 5.  | ClickAdd          | toadd                                            | thecommandsto |        |       | thenewtab                | orgroup. |     |     |     |
6. ClickImport/Export▶▶SavetheAuxilliaryApplicationUserInterface.The
thetoolkitribbonui.rbn
|     | modifiedlayoutissavedto |     |     |          |     |                              |     |     | filelocatedin | the |
| --- | ----------------------- | --- | --- | -------- | --- | ---------------------------- | --- | --- | ------------- | --- |
|     | textfolderwithin        |     |     | theVBAPI |     | applicationdirectory,thatis, |     |     |               |     |
<application_dir>/text
Note
The SavetheAuxilliaryApplicationUserInterfacebutton isenabled onlyif
|     | yousetthe |      | configurationoption |     |     | tk_enable_ribbon_custom_save |     |     |     |     |
| --- | --------- | ---- | ------------------- | --- | --- | ---------------------------- | --- | --- | --- | --- |
|     | to        | yes. |                     |     |     |                              |     |     |     |     |
7. ClickOK.
Note
|     | Youcannoteditthetoolkitribbonui.rbn |     |     |     |     |     |     | filemanually.                   |     |     |
| --- | ----------------------------------- | --- | --- | --- | --- | --- | --- | ------------------------------- | --- | --- |
| 70  |                                     |     |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |

| To Specify | the | Path | for | the | Ribbon |     | Definition |     | File |
| ---------- | --- | ---- | --- | --- | ------ | --- | ---------- | --- | ---- |
You canrenamethetoolkitribbonui.rbn toanotherfilename withthe
.rbnextension.
|     | ToenabletheVB |     |     | APIapplicationto |     |     | readthe | ribbondefinition |     |
| --- | ------------- | --- | --- | ---------------- | --- | --- | ------- | ---------------- | --- |
filehaving aname otherthantoolkitribbonui.rbn, itmustbe availableat
thelocation<application_dir>/text/ribbon. The methodintroduced
inthis sectionenablesyoutoloadthe ribbondefinitionfilefrom within aVB API
application.
methodIntroduced:
• IpfcSession.RibbonDefinitionfileLoad
| ThemethodIpfcSession.RibbonDefinitionfileLoad() |     |     |     |     |     |     |     |     | loadsa |
| ----------------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- | ------ |
specifiedribbon definitionfilefrom adefaultpathintotheCreo Parametric
| application.Theinput |     | argumentis |     | asfollows: |     |     |     |     |     |
| -------------------- | --- | ---------- | --- | ---------- | --- | --- | --- | --- | --- |
• file_name-Specifythename oftheribbon definitionfileincludingits
| extension.The |                      | defaultsearchpathfor |     |                       | thisfileis: |     |     |         |     |
| ------------- | -------------------- | -------------------- | --- | --------------------- | ----------- | --- | --- | ------- | --- |
| ○ The         | workingdirectoryfrom |                      |     | whereCreoParametricis |             |     |     | loaded. |     |
○ <application_text_dir>/ribbon
○ <application_text_dir>/language/ribbon
Note
◆ The locationofthe applicationtextdirectoryisspecifiedin theVB
API registryfile.
◆
AVB APIapplicationcanloadaribbon definitionfileonly once.After
|     | the applicationhasloadedtheribbon,callsmade |     |     |     |     |     |     | tothe | method      |
| --- | ------------------------------------------- | --- | --- | --- | --- | --- | --- | ----- | ----------- |
|     | IpfcSession.RibbonDefinitionfileLoad()      |     |     |     |     |     |     |       | toloadother |
ribbondefinitionfilesareignored.
| Loading    | Multiple |     | Applications |     |     | Using |     | the Ribbon |     |
| ---------- | -------- | --- | ------------ | --- | --- | ----- | --- | ---------- | --- |
| Definition | File     |     |              |     |     |       |     |            |     |
CreoParametricsupports loadingofmultiple .rbnfilesin the samesession.You
candevelopmultiple VBAPI applicationsthatshare thesametabs orgroups and
eachapplicationwillhaveitsownribbon definitionfile. Aseachapplicationis
loaded,its.rbnfilewillbe
|                         |     |     | read      | andapplied.Whenan |                   |     | applicationisunloaded, |          |     |
| ----------------------- | --- | --- | --------- | ----------------- | ----------------- | --- | ---------------------- | -------- | --- |
| thecontainersandcommand |     |     | createdby |                   | its.rbnfilewillbe |     |                        | removed. |     |
Forexample,considertwoVB APIapplications, namely, pt_geardesignand
pt_examplesthataddcommands tothe samegroupon atab onthe Ribbon
userinterface.The applicationpt_geardesignadds acommandPro/TOOLKIT
GearDesign tothe AdvancedModelinggroup on theModelingtabandthe
RibbonTabs,Groups,andMenuItems 71

applicationpt_examples addsacommand TKPartto the AdvancedModeling
groupon theModeltab. Theribbondefinitionfileforeachapplicationwillcontain
aninstructionto createtheAdvancedModelinggroupand ifboth theribbonfiles
areloaded,the groupwillbe createdonly onceand thetworibboncustomizations
| willbe merged | intothesamegroup. |     |     |     |     |     |
| ------------- | ----------------- | --- | --- | --- | --- | --- |
Thatis,ifboth theapplicationsarerunning in thesameCreoParametric session,
thenthe commands,Pro/TOOLKITGearDesign andTKPart willbeavailableunder
| theAdvancedModelinggroupon |     |     | theModel | tab. |     |     |
| -------------------------- | --- | --- | -------- | ---- | --- | --- |
Note
The orderin whichthecommands willbe displayedwithin thegroup will
| depend | onthe | orderof loadingofthe |     | .rbnfilefor |     | eachapplication. |
| ------ | ----- | -------------------- | --- | ----------- | --- | ---------------- |
Thefollowingimagedisplayscommandsaddedby twoVB APIapplicationsto
thesamegroup.
| Tosave thecustomizationwhenmultiple |     |     |     | applicationsare |     | loaded: |
| ----------------------------------- | --- | --- | --- | --------------- | --- | ------- |
1. ClickFile▶▶Options.The CreoParametricOptions dialogboxopens.
2. ClickCustomizeRibbon.
3. Inthe CustomizetheRibbonlist, selecta tabandcreateanewgroup in itor
| createanewtaband                                       |       | agroup        | init. |                          |          |     |
| ------------------------------------------------------ | ----- | ------------- | ----- | ------------------------ | -------- | --- |
| 4. Inthe Choosecommandsfromlist,selectTOOLKITCommands. |       |               |       |                          |          | The |
| commandscreatedby                                      |       | theVB         | API   | applicationaredisplayed. |          |     |
| 5. ClickAdd                                            | toadd | thecommandsto |       | thenewtab                | orgroup. |     |
6. ClickImport/Export▶▶SavetheAuxilliaryApplicationUserInterface.The
Save
| UICustomization |     | dialogbox | opens. |     |     |     |
| --------------- | --- | --------- | ------ | --- | --- | --- |
7. SelectaVB APIapplicationandClick Save.The modifiedlayoutissavedto
| the.rbnfileofthespecifiedVB |     |     |     | APIapplication. |                                 |     |
| --------------------------- | --- | --- | --- | --------------- | ------------------------------- | --- |
| 72                          |     |     |     |                 | Creo®ParametricVBAPIUser’sGuide |     |

TheSaveUICustomization dialogboxis shownin thefollowingimage:
| Localizing | the    | Ribbon | User         | Interface |
| ---------- | ------ | ------ | ------------ | --------- |
| Created    | by the | VB API | Applications |           |
Thelabels forthe customtabs, groups,andcascademenus belongingto theVB
APIapplicationcanbe translatedin thelanguagessupported byCreo Parametric.
Todisplaylocalizedlabels,specify thetranslated labelsin theribbonui.txt
| fileandsave | thisfileatthelocation<application_text_dir>/ |     |     |     |
| ----------- | -------------------------------------------- | --- | --- | --- |
<language>.For example,thetextfileforGermanlocalemustbe savedatthe
location<application_text_dir>/german/ribbonui.txt.
Createa filecontainingtranslationsfor eachofthe languagesin whichthe VB
APIapplicationislocalized.The Localizedtranslationfilesmustusethe UTF-8
encodingwith BOMcharacterfor thetranslated textto be displayedcorrectlyin
theuserinterface.
Theformatofthe ribbonui.txtfileisasshown below.Specifythe following
| linesfor eachlabel | entryinthe | file: |     |     |
| ------------------ | ---------- | ----- | --- | --- |
1. Ahashsign(#) followedby thelabel,as specifiedinthe ribbondefinitionfile.
2. Thelabel asspecifiedin theribbondefinitionfileandasdisplayedin the
ribbonuserinterface.
| 3. Thetranslated          | label. |                    |     |           |
| ------------------------- | ------ | ------------------ | --- | --------- |
| 4. Addanemptylineattheend |        | ofeachlabelentryin |     | the file. |
Forexample,ifthe CreoParametricapplicationcreatesatab withthe nameTK_
TAB havinga groupwith thename TK_GROUP,thenthetranslated filewill
containthefollowing:
#TK_TAB
RibbonTabs,Groups,andMenuItems 73

TK_TAB
| <translation | for | TK_TAB> |     |     |     |     |
| ------------ | --- | ------- | --- | --- | --- | --- |
<Empty_line>
#TK_GROUP
TK_GROUP
| <translation | for | TK_GROUP> |     |     |     |     |
| ------------ | --- | --------- | --- | --- | --- | --- |
<Empty_line>
| Support | for | Legacy |     | the VB | API |     |
| ------- | --- | ------ | --- | ------ | --- | --- |
Applications
Theuserinterfacefor CreoParametric 1.0hasbeenrestructuredto aribbon user
interface.This mayaffectthebehavior ofexisting theVB APIapplicationsthat
weredesignedto addcommands tospecific Pro/ENGINEERmenusor toolbars.
Thesemenusortoolbarsor bothhavebeenredesignedin Creo Parametric1.0.
| Thecommandsadded |     | bythe | theVB | APIapplicationsappearon |     | theCreo |
| ---------------- | --- | ----- | ----- | ----------------------- | --- | ------- |
Parametricribbonin theHometabunderthePro/TK group. Whenyouopena
| model,thePro/TK |     | group ison | theToolstab. |     |     |     |
| --------------- | --- | ---------- | ------------ | --- | --- | --- |
You canalsoarrangethe commandsaddedby thetheVB API applicationsundera
| newtaboran | existingtab |     | bycustomizing | theribbon | usingthe |     |
| ---------- | ----------- | --- | ------------- | --------- | -------- | --- |
CustomizeRibbon
tabin theCreoParametricOptions dialogbox. Fora listofall thecommands
| addedbythe                | VBAPI | applications,follow |     | thisprocedure: |                 |     |
| ------------------------- | ----- | ------------------- | --- | -------------- | --------------- | --- |
| 1. ClickFile▶▶Options.The |       |                     |     |                | dialogboxopens. |     |
CreoParametricOptions
2. ClickCustomizeRibbon.
3. Inthe Choosecommandsfromlist,selectTOOLKITCommands. Allthe
| commandsaddedby |     | legacyapplicationsarelisted. |     |     |     |     |
| --------------- | --- | ---------------------------- | --- | --- | --- | --- |
Note
Commandsthathavenotbeendesignated willnot havean iconorwillhavea
genericicon.
Refertothe CreoParametricHelp formoreinformation oncustomizing the
Ribbon.
| 74  |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

| Migration | of Legacy | VB API | Applications |
| --------- | --------- | ------ | ------------ |
Tomigrateexisting theVB APIapplicationstothe CreoParametricRibbon user
interfacewithoutcompilingthesourcecode:
1. LoadthetheVB API applicationsin Creo Parametricsothatthecommands
createdin theseapplicationsareavailablein theCustomizeRibbonuser
interface.
2. Modifythe ribbonuserinterfacelayoutandsavethe changestothe
toolkitribbonui.rbn.
| 3. Copythe | toolkitribbonui.rbn | to thelocation<application_ |     |
| ---------- | ------------------- | --------------------------- | --- |
text_dir>/ribbon.
4. ReloadtheVB APIapplicationorrestart Creo Parametric.The .rbnfilewill
| beloadedalong | withtheVB                | API applicationandthe          | commandswillbe |
| ------------- | ------------------------ | ------------------------------ | -------------- |
| visiblein     | theribbonuserinterfaceif | itsaccessibilitywillbevisible. |                |
RibbonTabs,Groups,andMenuItems 75

8
| Menus, | Commands, | and Pop-up |
| ------ | --------- | ---------- |
Menus
Introduction...............................................................................................................77
MenuBarDefinitions.................................................................................................77
MenusButtonsandMenus.........................................................................................77
DesignatingCommands.............................................................................................81
Pop-upMenus..........................................................................................................83
Thischapterdescribes themethodsprovided bythe VBAPI tocreateand modify
menus,buttons,andpop-upmenus inthe CreoParametricuserinterface.
Refertothe chapterRibbon Tabs, Groups,andMenu Itemson page67 for more
information.Also,refer to theCreo ParametricHelpfor moreinformationon
| customizingtheribbon | userinterface. |                                 |
| -------------------- | -------------- | ------------------------------- |
| 76                   |                | Creo®ParametricVBAPIUser’sGuide |

Introduction
TheVB APIclasses enableyou tosupplementthe CreoParametricribbonuser
interface.
OncetheVB API applicationis loaded,youcan addanewgroup to anexisting
| taborcreateanewtab |     | using | theCustomizeRibbontabin |     | the |     |
| ------------------ | --- | ----- | ----------------------- | --- | --- | --- |
CreoParametric
Optionsdialogboxin CreoParametric.Youwillnotbe abletomodifythe groups
thatare definedbyCreo Parametric.Ifthetranslated messagesareavailablefor
thenewlyaddedtabs orgroups,thenCreo Parametricwillusethem bysearching
| forthesamestring |     | inthe listof | stingbasedmessagesloaded. |     |     |     |
| ---------------- | --- | ------------ | ------------------------- | --- | --- | --- |
You cancustomizetheribbonuserinterfaceonly for aparticular modein Creo
Parametric.For example,ifyoucustomizetheribbonuser interfaceandsaveitto
| thetoolkitribbonui.rbn |     |     | fileinthe | Partmode,thenonloading |     | Creo |
| ---------------------- | --- | --- | --------- | ---------------------- | --- | ---- |
Parametricthecustomizeduserinterfacewillbe visibleonlyin thePart mode.To
viewaparticulartaborgroupin allthe modes,youmustcustomizethe ribbon
userinterfaceandsavethe toolkitribbonui.rbn filein eachmode.Refer
tothe CreoParametricFundamentalsHelp formoreinformation oncustomizing
theribbon.
| Menu | Bar | Definitions |     |     |     |     |
| ---- | --- | ----------- | --- | --- | --- | --- |
• Menu—Amenu,suchasthe Filemenu,ora sub-menu,suchasthe Manage
| Filemenuunderthe |     | Filemenu. |     |     |     |     |
| ---------------- | --- | --------- | --- | --- | --- | --- |
• Menu button—Anameditem ina groupor menuthatisused tolaunchasetof
instructions.
| • Pop-upmenu—Amenuinvokedby |     |     |     | selectionof | anitem | in theCreo |
| --------------------------- | --- | --- | --- | ----------- | ------ | ---------- |
Parametricgraphicswindow.
• Command—Aprocedurein CreoParametricthatmaybe activatedfroma
button.
| Menus | Buttons |     | and | Menus |     |     |
| ----- | ------- | --- | --- | ----- | --- | --- |
Thefollowingmethodsenableyouto addnewmenubuttons in anylocationon
| theRibbon | userinterface. |     |     |     |     |     |
| --------- | -------------- | --- | --- | --- | --- | --- |
MethodsIntroduced:
• IpfcSession.UICreateCommand()
• IpfcSession.UICreateMaxPriorityCommand()
• IpfcUICommandActionListener.OnCommand()
Menus,Commands,andPop-upMenus 77

ThemethodIpfcSession.UICreateCommand() createsa
| IpfcUICommandobjectthatcontains |     |     | a                     |
| ------------------------------- | --- | --- | --------------------- |
| IpfcUICommandActionListener.    |     |     | You shouldoverridethe |
IpfcUICommandActionListener.OnCommand()method withthe code
| thatyouwant | to executewhen | theuserclicks | abutton. |
| ----------- | -------------- | ------------- | -------- |
ThemethodIpfcSession.UICreateMaxPriorityCommand() createsa
IpfcUICommandobjecthaving maximumcommandpriority. Thepriority ofthe
actionrefers to thelevel ofprecedencetheaddedactiontakesover otherCreo
Parametricactions.Maximum priorityactionsdismissall otheractionsexcept
asynchronousactions.
Maximumcommandpriorityshouldbe usedonly in commandsthatopen and
activateanewmodelin awindow. Createallothercommands usingthe method
IpfcSession.UICreateCommand().
IpfcUICommandListener.OnCommand()is
Thelistenermethod called
whenthecommand isactivatedinCreo Parametricby pressinga button.
Designatethe commandusing thefunction IpfcUICommand.Designate()
andaddabutton to theribbonuserinterfaceusing theusing theCustomizeRibbon
tabin theCreoParametricOptions dialogbox. Thisoperationbinds thecommand
tothe button.
| Finding | Creo Parametric |     | Commands |
| ------- | --------------- | --- | -------- |
Thismethodenablesyouto find existingCreo Parametriccommandsin orderto
| modifytheir | behavior. |     |     |
| ----------- | --------- | --- | --- |
MethodIntroduced:
• IpfcSession.UIGetCommand()
ThemethodIpfcSession.UIGetCommand() IpfcUICommand
returnsa
objectrepresentingan existing CreoParametriccommand.The methodallows
youtofind thecommandID foran existingcommand sothatyoucanaddan
accessfunctionor bracketfunctionto thecommand.Youmustknow thenameof
| thecommandin | order tofind | itsID. |     |
| ------------ | ------------ | ------ | --- |
Tofind thenameof an actioncommand,click thecorrespondingiconon the
ribbonuserinterfaceandthencheckthe lastentryin thetrail file.For example,for
| theSaveicon, | thetrail filewillhavethecorrespondingentry: |     |     |
| ------------ | ------------------------------------------- | --- | --- |
| ~ Command    | `ProCmdModelSave`                           |     |     |
Theactionnamefor theSaveiconisProCmdModelSave. This stringcanbe
usedasinputto IpfcSession.UIGetCommand() to getthecommand ID.
78 Creo®ParametricVBAPIUser’sGuide

You candetermineacommandID stringforan option withoutaniconby
searchingthrough theresourcefileslocated inthe <creo_loadpoint>\
<datecode>\CommonFiles\text\resource directory.Ifyousearchfor
themenubutton name,thelinewillcontainthe correspondingactioncommand
forthebutton.
| Access | Listeners | for | Commands |     |     |     |     |
| ------ | --------- | --- | -------- | --- | --- | --- | --- |
Thesemethodsallowyouto applyan accesslistener toa command.Theaccess
listenerdetermineswhetherornotthe commandisvisibleatthecurrenttime in
thecurrentsession.
MethodsIntroduced:
• IpfcActionSource.AddActionListener()
• IpfcUICommandAccessListener.OnCommandAccess()
Usethe methodIpfcActionSource.AddActionListener() to registera
| newIpfcUICommandAccessListeneron |     |     |     | anycommand(created |     |     | eitherby |
| -------------------------------- | --- | --- | --- | ------------------ | --- | --- | -------- |
anapplicationorCreo Parametric).Thislistener willbecalled whenbuttons based
| onthecommand | mightbeshown. |     |     |     |     |     |     |
| ------------ | ------------- | --- | --- | --- | --- | --- | --- |
Thelistenermethod
IpfcUICommandAccessListener.OnCommandAccess()
allowsyouto
imposean accessfunction ona particularcommand.The methoddeterminesif the
| actionor             | commandshouldbe       | available,unavailable, |                      |     | orhidden. |     |     |
| -------------------- | --------------------- | ---------------------- | -------------------- | --- | --------- | --- | --- |
| Thepotential         | returnvaluesarelisted |                        | in theenumeratedtype |     |           |     |     |
| EpfcCommandAccessand |                       | areasfollows:          |                      |     |           |     |     |
• EpfcACCESS_REMOVE—The buttonis notvisibleandifall ofthe menu
| buttonsin | the containingmenu |     | possessan | accessfunction |     | returning |     |
| --------- | ------------------ | --- | --------- | -------------- | --- | --------- | --- |
EpfcACCESS_REMOVE,the containingmenuwillalsobe removedfrom the
CreoParametricuserinterface..
| • EpfcACCESS_INVISIBLE—Thebutton |     |     |     | isnotvisible. |     |     |     |
| -------------------------------- | --- | --- | --- | ------------- | --- | --- | --- |
• EpfcACCESS_UNAVAILABLE—Thebutton isvisible,butgray andcannot
beselected.
| • EpfcACCESS_DISALLOW—The |     |                         | buttonshowsas |     | available,butthe |     |     |
| ------------------------- | --- | ----------------------- | ------------- | --- | ---------------- | --- | --- |
| commandwillnotbe          |     | executedwhenitischosen. |               |     |                  |     |     |
• EpfcACCESS_AVAILABLE—Thebutton isnotgray andcanbe selectedby
| theuser. | Thisis thedefaultvalue. |     |     |     |     |     |     |
| -------- | ----------------------- | --- | --- | --- | --- | --- | --- |
Menus,Commands,andPop-upMenus 79

| Example         | 1: Command                               | Access | Listeners |     |     |
| --------------- | ---------------------------------------- | ------ | --------- | --- | --- |
| Thesamplecodein | thefilepfcCommandExamples.vb<creo_vbapi_ |        |           |     |     |
loadpoint>/vbapi_examplesdemonstratesthe usageof theaccesslistener
methodfora particularcommand.The OnCommandAccess functionreturns
| EpfcACCESS_UNAVAILABLEthatdisablesbutton |     |     |     | associatedwith | the |
| ---------------------------------------- | --- | --- | --- | -------------- | --- |
command,if themodelis notpresentor ifitis notoftypePART.Else, the
functionreturnsEpfcACCESS_AVAILABLE thatenablesbutton associatedwith
thecommand.
| Bracket | Listeners | for Commands |     |     |     |
| ------- | --------- | ------------ | --- | --- | --- |
Thesemethodsallowyouto applya bracketlistenerto acommand.The bracket
listeneriscalledbefore andafterthecommandruns, whichallowsyour
applicationto providecustomlogicto executewhenever thecommandis selected.
MethodsIntroduced:
• IpfcActionSource.AddActionListener()
• IpfcUICommandBracketListener.OnBeforeCommand()
• IpfcICommandBracketListener.OnAfterCommand()
Usethe methodIpfcActionSource.AddActionListener() to registera
newIpfcUICommandBracketListener onanycommand (createdeitherby
anapplicationorCreo Parametric.Thislistener willbe calledwhen thecommand
| isselectedbythe | user. |     |     |     |     |
| --------------- | ----- | --- | --- | --- | --- |
Thelistenermethods
| IpfcICommandBracketListener.OnBeforeCommand() |     |     |     |     | and      |
| --------------------------------------------- | --- | --- | --- | --- | -------- |
| IpfcUICommandBracketListener.OnAfterCommand() |     |     |     |     | allowthe |
creationoffunctions thatwillbecalled immediatelybefore andafterexecutionof
agivencommand.Thesemethods areusedto addthebusiness logicto thestartor
| end(orboth)of | anexisting | CreoParametriccommand. |     |     |     |
| ------------- | ---------- | ---------------------- | --- | --- | --- |
ThemethodIpfcUICommandBracketListener.OnBeforeCommand()
| couldalsobe | usedto cancelan | upcomingcommand.Todothis,throw |     |     | a   |
| ----------- | --------------- | ------------------------------ | --- | --- | --- |
IpfcXCancelProEAction exceptionfrom thebody ofthelistener method
usingCCpfcXCancelProEAction.Throw().
| Example                                        | 2: Bracket                   | Listeners |     |                 |     |
| ---------------------------------------------- | ---------------------------- | --------- | --- | --------------- | --- |
| Thesamplecodein                                | thefilepfcCommandExamples.vb |           |     | locatedat<creo_ |     |
| vbapi_loadpoint>/vbapi_examplesdemonstratesthe |                              |           |     | usageofthe      |     |
bracketlistenermethodsthatarecalled beforeandafterwhenthe usertriesto
renamethemodel. Ifthemodelcontainsa certainparameter, therenameattempt
| willbe abortedby | thislistener. |     |     |                                 |     |
| ---------------- | ------------- | --- | --- | ------------------------------- | --- |
| 80               |               |     |     | Creo®ParametricVBAPIUser’sGuide |     |

| Designating |     |                         | Commands |     |                          |     |     |
| ----------- | --- | ----------------------- | -------- | --- | ------------------------ | --- | --- |
| UsingtheVB  |     | API youcandesignateCreo |          |     | Parametriccommands.These |     |     |
commandscanlaterappear in theCreo Parametricribbonuserinterface.
| Toadda | commandyoumust: |     |     |     |     |     |     |
| ------ | --------------- | --- | --- | --- | --- | --- | --- |
1. Defineor addthecommandto be initiatedonclickingthe iconin theVB API
application.
| 2.  | Optionallydesignateaniconbutton |     |     |     | to beused | withthecommand. |     |
| --- | ------------------------------- | --- | --- | --- | --------- | --------------- | --- |
3. Designatethecommandto appear inthe CustomizeRibbontabin theCreo
ParametricOptionsdialogbox.
Note
Referto thechapter onRibbon Tabs, Groups,and MenuItemson page67
for moreinformation.Also, referto theCreo ParametricParameticHelp
|     | for | moreinformationon |     | customizingthe |     | RibbonUserInterface. |     |
| --- | --- | ----------------- | --- | -------------- | --- | -------------------- | --- |
4. Savetheconfigurationin Creo Parametricsothatchangestothe ribbonuser
|         | interfaceappear |       | whenanewsessionof |     | CreoParametricis |     | started. |
| ------- | --------------- | ----- | ----------------- | --- | ---------------- | --- | -------- |
| Command |                 | Icons |                   |     |                  |     |          |
MethodIntroduced:
• IpfcUICommand.SetIcon()
ThemethodIpfcUICommand.SetIcon() allowsyouto designateaniconto
beused withthecommand youcreated.Themethod addstheiconto theCreo
Parametriccommand.Specifythe nameofthe iconfile,includingthe extensionas
theinputargumentfor thismethod.Avalidformatforthe iconfileis astandard
.GIF,.JPG, or.PNG.PTCrecommends using.PNGformat. Alliconsin the
CreoParametric ribbonareeither16x16(small)or32x32(large)size. The
| namingconventionfor |     |     | theiconsis | asfollows: |     |     |     |
| ------------------- | --- | --- | ---------- | ---------- | --- | --- | --- |
• Smallicon—<iconname.ext>
| •   | Large icon—<iconname_large.ext> |     |     |     |     |     |     |
| --- | ------------------------------- | --- | --- | --- | --- | --- | --- |
Menus,Commands,andPop-upMenus 81

Note
• Thelegacynaming conventionfor icons<icon_name_icon_size.ext>
willnotbesupported in futurereleasesofCreo Parametric.Theiconsizewas
addedasa suffixto thename oftheicon. Forexample,the legacynaming
conventionfor small iconswas iconname16X16.ext. Itis recommended
tousethe standardnamingconventionsforicons, thatis,iconname.extor
iconname_large.ext.
• Whilespecifyingthename oftheiconfile, donotspecify thefull pathtothe
iconnames.
| Theapplicationsearchesforthe         |     | iconfilesin | thefollowinglocations: |             |     |
| ------------------------------------ | --- | ----------- | ---------------------- | ----------- | --- |
| • <creo_loadpoint>\<datecode>\Common |     |             |                        | Files\text\ |     |
resource
• <Applicationtextdir>\resource
• <Applicationtextdir>\<language>\resource
Thelocationof theapplicationtextdirectoryis specifiedinthe registryfile.
Commandsthatdonot havean iconassignedto themdisplaythebutton label.
You mayalso usethismethodto assignasmall icontoa button.The iconappears
| tothe leftof | thebutton | label.      |     |     |     |
| ------------ | --------- | ----------- | --- | --- | --- |
| Designating  |           | the Command |     |     |     |
MethodIntroduced:
• IpfcUICommand.Designate()
Thismethodallows youdesignatethecommandas availablein the Customize
| Ribbontab | inthe |     | dialogofCreo | Parametric.After | aVB |
| --------- | ----- | --- | ------------ | ---------------- | --- |
CreoParametricOptions
| APIapplicationhasused |     | themethodIpfcUICommand.Designate() |     |     | ona |
| --------------------- | --- | ---------------------------------- | --- | --- | --- |
command,can addthebutton associatedwith thiscommand intothe Creo
Parametricribbonuserinterface.
Ifthismethodis notcalled,thebutton willnot bevisiblein the ToolkitCommands
listin theCustomizeRibbontab in theCreoParametricOptions dialogofCreo
Parametric.
| Theargumentsto | thismethodare: |     |     |     |     |
| -------------- | -------------- | --- | --- | --- | --- |
• Label—Themessagestringthatrefersto theiconlabel.This label (storedin
themessagefile)identifies thetextseenwhen thebutton isdisplayed. Ifthe
| 82  |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | ------------------------------- | --- | --- |

commandisnot assignedanicon, thebutton labelstringappears on thetoolbar
| buttonby | default. |     |     |     |     |
| -------- | -------- | --- | --- | --- | --- |
• Help—Theone-line Helpfor theicon. Thislabel (storedin themessagefile)
| identifiesthe | helplineseen |     | whenthemouse | movesoverthe | icon. |
| ------------- | ------------ | --- | ------------ | ------------ | ----- |
• Description—Themessageappears inthe CustomizeRibbontabin theCreo
ParametricOptionsdialogboxand alsowhenDescription isclicked inCreo
Parametric.
• MessageFile—Themessagefilename.Allthelabels includingtheone-line
| Helplabels | mustbe | presentin | themessagefile. |     |     |
| ---------- | ------ | --------- | --------------- | --- | --- |
Note
| This | filemustbe | in  | thedirectory<text_path>/text |     | or<text_ |
| ---- | ---------- | --- | ---------------------------- | --- | -------- |
path>/text/<language>.
| Placing | the Button |     |     |     |     |
| ------- | ---------- | --- | --- | --- | --- |
Oncethebutton hasbeencreatedusing themethodsdiscussed, place thebutton on
theCreoParametricribbon userinterface.Refertothe chapteron RibbonTabs,
Groups,andMenu Itemson page67 for moreinformation.Also,refer to theCreo
ParametricHelpfor moreinformationon customizingthe RibbonUserInterface.
| Pop-up | Menus |     |     |     |     |
| ------ | ----- | --- | --- | --- | --- |
CreoParametricprovidesshortcutmenusthatcontain frequentlyusedcommands
appropriateto thecurrentlyselecteditems.You canaccessashortcut menuby
| right-clickingaselecteditem.Shortcut |     |     | menusareaccessiblein: |     |     |
| ------------------------------------ | --- | --- | --------------------- | --- | --- |
• Graphicswindow
• ModelTree
• Somedialogboxes
• Anyareawhereyoucan performanobject-action operationbyselecting an
| itemandchoosing |     | acommandto | performonthe | selecteditem. |     |
| --------------- | --- | ---------- | ------------ | ------------- | --- |
Themethodsdescribed in thissectionallowyouto addmenus toa graphics
windowpop-upmenu.
Menus,Commands,andPop-upMenus 83

| Adding | a Pop-up | Menu | to the Graphics |     | Window |
| ------ | -------- | ---- | --------------- | --- | ------ |
You canactivatedifferent pop-upmenusduringa givensessionofCreo
Parametric.Everytime theCreo Parametriccontextchangeswhen youopena
differentmodeltype,enterdifferent toolsorspecialmodes suchasEdit, adifferent
pop-upmenuiscreated.When CreoParametricmovesto the nextcontext,the
pop-upmenumaybedestroyed.
Asaresultofthis, theVB APIapplicationsmustattachabutton tothe pop-up
menuduringinitializationofthepop-upmenu.The VB APIapplicationisnotified
eachtime aparticular pop-upmenuis created,whichthenallowstheuserto add
tothe pop-upmenu.
Usethe followingprocedureto additemsto pop-upmenus inthe graphics
window:
1. Obtainthe nameof theexistingpop-up menusto whichyouwanttoaddanew
| menuusing            | thetrailfile. |                        |     |     |     |
| -------------------- | ------------- | ---------------------- | --- | --- | --- |
| 2. Createcommandsfor |               | thenewpop-upmenuitems. |     |     |     |
3. Implementaccesslisteners toprovidevisibilityinformationfor theitems.
4. Addanactionlistener to thesessionto listenfor pop-upmenuinitialization.
5. Inthe listenermethod,ifthe pop-upmenuisthe correctmenuto whichyou
| wishto | addthebutton,thenaddit. |     |     |     |     |
| ------ | ----------------------- | --- | --- | --- | --- |
Thefollowingsectionsdescribeeachofthesesteps in detail.You canaddpush
buttonsandcascade menusto thepop-upmenus.Youcan addpop-upmenuitems
onlyin theactivewindow.Youcannotusethisprocedure to removeitemsfrom
existingmenus.
| Using | the Trail | File | to Determine | Existing | Pop-up |
| ----- | --------- | ---- | ------------ | -------- | ------ |
| Menu  | Names     |      |              |          |        |
Thetrailfilein CreoParametriccontains acomment thatidentifiesthename of
thepop-upmenuif theconfigurationoption,auxapp_popup_menu_info is
setto yes.
Forexample,thepop-upmenu,EditProperties, hasthefollowingcommentin the
trailfile:
| ~ Close    | `rmb_popup` | `PopupMenu`      |     |     |     |
| ---------- | ----------- | ---------------- | --- | --- | --- |
| ~ Activate | `rmb_popup` | `EditProperties` |     |     |     |
!Command ProCmdEditPropertiesDtm was pushed from the software.
| !Item was | selected | from   | popup menu 'popup_mnu_edit' |     |     |
| --------- | -------- | ------ | --------------------------- | --- | --- |
| Listening | for      | Pop-up | Menu Initialization         |     |     |
MethodsIntroduced:
| 84  |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

• IpfcActionSource.AddActionListener()
• IpfcActionSource.AddActionListenerWithType()
• IpfcPopupmenuListener.OnPopupmenuCreate()
| Usethe | methodIpfcActionSource.AddActionListener() |     |     |     |     |     |     | or  |
| ------ | ------------------------------------------ | --- | --- | --- | --- | --- | --- | --- |
IpfcActionSource.AddActionListenerWithType()to
registera
newIpfcPopupmenuListener tothe session.This listener willbe called
| whenpop-upmenusare                                 |     | initialized. |     |     |     |     |     |     |
| -------------------------------------------------- | --- | ------------ | --- | --- | --- | --- | --- | --- |
| ThemethodIpfcPopupmenuListener.OnPopupmenuCreate() |     |              |     |     |     |     |     | is  |
calledafterthepop-upmenuis createdinternallyin CreoParametricandmaybe
| usedto assignapplication-ownedbuttons |     |            |     |       | to  | thepop-upmenu. |     |     |
| ------------------------------------- | --- | ---------- | --- | ----- | --- | -------------- | --- | --- |
| Accessing                             |     | the Pop-up |     | Menus |     |                |     |     |
Thepropertydescribed in thissectionprovidesthename ofthepop-up menus
| usedto accessthesemenuswhileusing |     |     |     |     | othermethods. |     |     |     |
| --------------------------------- | --- | --- | --- | --- | ------------- | --- | --- | --- |
PropertyIntroduced:
• IpfcPopupmenu.Name
ThepropertyIpfcPopupmenu.Name
|        |         |     |     |        | returnsthename |       | ofthepop-up | menu. |
| ------ | ------- | --- | --- | ------ | -------------- | ----- | ----------- | ----- |
| Adding | Content | to  | the | Pop-up |                | Menus |             |       |
MethodsIntroduced:
• IpfcPopupmenu.AddButton()
• IpfcPopupmenu.AddMenu()
Usethe methodIpfcPopupmenu.AddButton()to addanewitem to apop-
| upmenu.The             | inputarguments |     | are:                  |     |     |     |                |     |
| ---------------------- | -------------- | --- | --------------------- | --- | --- | --- | -------------- | --- |
| • Command—Specifiesthe |                |     | commandassociatedwith |     |     |     | thepop-upmenu. |     |
• Options– AIpfcPopupmenuOptionsobjectcontainingotheroptions for
| themethod.The |     | optionsthatmaybe |     |     | includedare: |     |     |     |
| ------------- | --- | ---------------- | --- | --- | ------------ | --- | --- | --- |
○ PositionIndex—Specifies thepositionin thepop-up menuatwhich
to addthemenubutton. Pass nullto addthe buttonto thebottom ofthe
| menu.Use |                | thepropertyIpfcPopupmenuOptions.PositionIndex |     |     |     |     |     |     |
| -------- | -------------- | --------------------------------------------- | --- | --- | --- | --- | --- | --- |
| to       | setthisoption. |                                               |     |     |     |     |     |     |
○ Name—Specifies thenameof theaddedbutton. Thebutton nameis placed
in thetrailfilewhen theuserselectsthemenubutton. Usethe property
| IpfcPopupmenuOptions.Nameto   |     |     |     |     |     | setthisoption. |     |     |
| ----------------------------- | --- | --- | --- | --- | --- | -------------- | --- | --- |
| Menus,Commands,andPop-upMenus |     |     |     |     |     |                |     | 85  |

○ SetLabel—Specifiesthebutton label.This labelidentifies thetext
| displayedwhenthe           |     | buttonis displayed.Use | theproperty       |
| -------------------------- | --- | ---------------------- | ----------------- |
| IpfcPopupmenuOptions.Label |     |                        | to setthisoption. |
○ Helptext—Specifiesthehelpmessageassociatedwith thebutton. Use
thepropertyIpfcPopupmenuOptions.Helptext to setthisoption.
Usethe methodIpfcPopupmenu.AddMenu() to addanewcascademenu to
| anexisting | pop-upmenu. |     |     |
| ---------- | ----------- | --- | --- |
Theargumentfor thismethodis aIpfcPopupmenuOptionsobject,whose
membershavethesamepurpose asdescribedfornewlyadded buttons.This
methodreturnsanewIpfcPopupmenu
objectto whichyoumayaddnew
buttons.
| Example | 3: Creating | a Pop-up | Menu |
| ------- | ----------- | -------- | ---- |
Thesamplecodein thefilepfcPopupMenuExamples.vblocatedat<creo_
vbapi_loadpoint>/vbapi_examplesdemonstratesthe usageofUI
| functionsto | addanewmodeltreepop-upmenu. |     |                                 |
| ----------- | --------------------------- | --- | ------------------------------- |
| 86          |                             |     | Creo®ParametricVBAPIUser’sGuide |

9
Models
OverviewofModelObjects.........................................................................................88
GettingaModelObject..............................................................................................88
ModelDescriptors.....................................................................................................88
RetrievingModels.....................................................................................................89
ModelInformation.....................................................................................................91
ModelOperations......................................................................................................94
RunningCreoModelCHECK......................................................................................96
Thischapterdescribes howtoprogramon themodellevel usingthe VBAPI.
87

| Overview |     | of Model |     | Objects |     |     |
| -------- | --- | -------- | --- | ------- | --- | --- |
Modelscanbe anyCreo Parametricfiletype,includingparts, assemblies,
drawings,sections,andnotebook.The classesin themodulepfcModel provide
genericaccessto models,regardless oftheirtype. Theavailablemethodsenable
| youtodo                    | thefollowing: |       |         |     |     |     |
| -------------------------- | ------------- | ----- | ------- | --- | --- | --- |
| • Accessinformationabouta  |               |       | model.  |     |     |     |
| • Open,copy,rename,andsave |               |       | amodel. |     |     |     |
| Getting                    | a             | Model | Object  |     |     |     |
MethodsandPropertiesIntroduced:
• IpfcFamilyTableRow.CreateInstance()
• IpfcSelection.SelModel
• IpfcBaseSession.GetModel()
• Ipfc.BaseSession.CurrentModel
• IpfcBaseSession.GetActiveModel()
• Ipfc.BaseSession.ListModels()
• IpfcBaseSession.GetByRelationId()
• IpfcWindow.Model
| Thesemethodsgeta                  |     | modelobjectthatis |     | alreadyin  | session. |              |
| --------------------------------- | --- | ----------------- | --- | ---------- | -------- | ------------ |
| ThepropertyIpfcSelection.SelModel |     |                   |     | returnsthe |          | modelthatwas |
interactivelyselected.
ThemethodIpfcBaseSession.GetModel() returnsa modelbasedon its
nameandtype, whereasIpfcBaseSession.GetByRelationId() returns
| amodelinan | assembly | thathasthe | specifiedintegeridentifier. |     |     |     |
| ---------- | -------- | ---------- | --------------------------- | --- | --- | --- |
ThepropertyIpfcBaseSession.CurrentModel returnsthe currentactive
model.
| ThemethodIpfcBaseSession.GetActiveModel() |     |     |     |     |     | returns theactive |
| ----------------------------------------- | --- | --- | --- | --- | --- | ----------------- |
CreoParametricmodel.
Usethe methodIpfcBaseSession.ListModels() toreturn asequenceof
| allthemodelsin |     | session. |     |     |     |     |
| -------------- | --- | -------- | --- | --- | --- | --- |
Formoremethodsthatreturn solidmodels, refer tothe chapterSolidon page154.
| Model | Descriptors |     |     |     |     |     |
| ----- | ----------- | --- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
| 88  |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

• CCpfcModelDescriptor.Create()
• IpfcModelDescriptor_CreateFromFileName()
• IpfcModelDescriptor.SetGenericName
• IpfcModelDescriptor.SetInstanceName
• IpfcModelDescriptor.SetType
• IpfcModelDescriptor.SetDevice
• IpfcModelDescriptor.SetPath
• IpfcModelDescriptor.FileVersion
• IpfcModelDescriptor.GetFullName
• IpfcModel.GetFullName
Modeldescriptorsaredataobjectsusedto describeamodelfileanditslocationin
thesystem.The methodsin themodeldescriptorenableyoutosetspecific
informationthatenablesCreoParametricto find thespecificmodelyouwant.
Thestaticutilitymethod CCpfcModelDescriptor.Create()allowsyouto
specifyas datato be enteredamodeltype, an instancename,anda genericname.
Themodeldescriptorconstructsthefullname ofthe modelas astring,asfollows:
String FullName = InstanceName+"<"+GenericName+">";
// As long as the
// generic name is
// not an empty
// string ("")
Ifyouwantto loadamodelthatisnota familytableinstance, passanempty
stringasthegenericname argument sothatthefull nameof themodelis
constructedcorrectly. Ifthemodelisa familytableinterface,youshould specify
boththeinstance andgenericnames.
Note
Youareallowedto setotherfieldsin the modeldescriptorobject,buttheymay
be ignoredbysomemethods.
Thestaticutilitymethod
IpfcModelDescriptor.CreateFromFileName() allowsyouto createa
newmodeldescriptorfrom agivena filename. Thefilenameis astringin the
form<name>.<extension>.
Retrieving Models
MethodsIntroduced:
Models 89

• IpfcBaseSession.RetrieveModel()
• IpfcBaseSession.RetrieveModelWithOpts()
• IpfcBaseSession.OpenFile()
• IpfcSolid.HasRetrievalErrors()
ThesemethodscauseCreo Parametrictoretrieve themodelthatcorresponds to
theIpfcModelDescriptor argument.
ThemethodIpfcBaseSession.RetrieveModel() retrieves thespecified
modelintotheCreo Parametricsessiongivenitsmodeldescriptorfrom astandard
directory. This methodignoresthe pathargumentspecifiedin themodel
descriptor.But thismethoddoesnot createawindowfor it,nordoesitdisplaythe
modelanywhere.
| ThemethodIpfcBaseSession.RetrieveModelWithOpts() |     |     | retrieves |     |
| ------------------------------------------------ | --- | --- | --------- | --- |
thespecifiedmodelintothe CreoParametricsessionbasedon thepathspecified
bythemodeldescriptor. Thepathcan bea diskpath,a workspacepath, ora
| commonspacepath. | The Optsargument(given | by the |     |     |
| ---------------- | ---------------------- | ------ | --- | --- |
IpfcRetrieveModelOptionsobject)providesthe userwiththe optionto
specifysimplifiedrepresentations.
ThemethodIpfcBaseSession.OpenFile() bringsthemodelinto
memory, opensa newwindowforit(or usesthebasewindow,if itisempty),and
| displaysthe model. |     |     |     |     |
| ------------------ | --- | --- | --- | --- |
Note
| IpfcBaseSession.OpenFile() |     | actuallyreturnsa | handleto the |     |
| -------------------------- | --- | ---------------- | ------------ | --- |
windowithascreated.
Thefileversionsetby theIpfcModelDescriptor.FileVersion property
| isignoredbythe | IpfcBaseSession.RetrieveModelWithOpts() |     |     | and |
| -------------- | --------------------------------------- | --- | --- | --- |
IpfcBaseSession.OpenFile()methods.Instead, thelatestfileversionof
| themodelisusedby | thesetwomethods. |     |     |     |
| ---------------- | ---------------- | --- | --- | --- |
Toget ahandletothe modelyouneed,usethepropertyIpfcWindow.Model.
ThemethodIpfcSolid.HasRetrievalErrors() returns atrue valueif
thefeaturesin thesolidmodelweresuppressedduringthe RetrieveModelor
OpenFileoperations.This methodmustbecalledimmediatelyafterthe
IpfcBaseSession.RetrieveModel()methodor anequivalent retrieval
method.
| 90  |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | ------------------------------- | --- | --- |

Example Code: Retrieving a Model
Thesamplecodein thefilepfcModelsExamples.vblocated at<creo_
vbapi_loadpoint>/vbapi_examplesdemonstrateshowto retrieve a
model.
Model Information
MethodsandPropertiesIntroduced:
• IpfcModel.FileName
• IpfcModel.CommonName
• IpfcModel.IsCommonNameModifiable()
• IpfcModel.FullName
• IpfcModel.GenericName
• IpfcModel.InstanceName
• IpfcModel.Origin
• IpfcModel.RelationId
• IpfcModel.Descr
• IpfcModel.Type
• IpfcModel.IsModified
• IpfcModel.Version
• IpfcModel.Revision
• IpfcModel.Branch
• IpfcModel.ReleaseLevel
• IpfcModel.VersionStamp
• IpfcModel.ListDependencies()
• IpfcModel.CleanupDependencies()
• IpfcModel.ListDeclaredModels()
• IpfcModel.CheckIsModifiable()
• IpfcModel.CheckIsSaveAllowed()
ThepropertyIpfcModel.FileName retrievesthe modelfilename inthe
"name"."type"format.
ThepropertyIpfcModel.CommonName retrievesthe commonname forthe
model.This nameisdisplayedfor themodelin WindchillPDMLink.
Models 91

Usethe methodIpfcModel.IsCommonNameModifiable() to identifyif
thecommon nameofthe modelcan be modified.Youcan modifythenamefor
modelsthatarenot yetownedbyWindchillPDMLink, orin certainsituationsif
theconfigurationoptionlet_proe_rename_pdm_objects issetto yes.
ThepropertyIpfcModel.FullName retrievesthe fullname ofthemodelin
| theinstance<generic> |     | format. |     |     |     |     |
| -------------------- | --- | ------- | --- | --- | --- | --- |
ThepropertyIpfcModel.GenericName retrieves thename ofthegeneric
mustbeNULLor
| model.Ifthe | modelis notan | instance,thisname |     |     | anempty |     |
| ----------- | ------------- | ----------------- | --- | --- | ------- | --- |
string.
ThepropertyIpfcModel.InstanceName retrieves thenameof themodel.If
| themodelisan | instance,this | methodretrievesthe | instancename. |     |     |     |
| ------------ | ------------- | ------------------ | ------------- | --- | --- | --- |
ThepropertyIpfcModel.Origin returnsthecomplete pathtothe filefrom
| whichthemodelwasopened.   |     | Thispathcan | bea locationondiskfrom |     | a   |     |
| ------------------------- | --- | ----------- | ---------------------- | --- | --- | --- |
| Windchillworkspace,orfrom |     | adownloaded | URL.                   |     |     |     |
ThepropertyIpfcModel.RelationId
|                 |       |          | retrievesthe | relationidentifierof |     | the |
| --------------- | ----- | -------- | ------------ | -------------------- | --- | --- |
| specifiedmodel. | Itcan | be NULL. |              |                      |     |     |
ThepropertyIpfcModel.Descr returnsthe descriptorforthe specifiedmodel.
Modeldescriptorscanbe usedto representmodelsnotcurrentlyin session.
Note
| From                                      | Pro/ENGINEERWildfire4.0onwards,theproperties |                   |                        |     |     |     |
| ----------------------------------------- | -------------------------------------------- | ----------------- | ---------------------- | --- | --- | --- |
| IpfcModel.FullName,IpfcModel.GenericName, |                                              |                   |                        |     | and |     |
| IpfcModel.Descr                           |                                              | throw anexception | IpfcXtoolkitCantOpenif |     |     |     |
calledon amodelinstancewhose immediategenericisnot insession. Handle
thisexception andtypecast themodelasIpfcSolid, whichinturn canbe
asIpfcFamilyMember,andusethe
| typecast                                   |     |     | method |     |           |     |
| ------------------------------------------ | --- | --- | ------ | --- | --------- | --- |
| IpfcFamilyMember.GetImmediateGenericInfo() |     |     |        |     | toget the |     |
modeldescriptoroftheimmediategenericmodel. Themodeldescriptorcan be
usedto derive thefull nameorgenericname ofthemodel. Ifyouwishto
switchoffthisbehavior andcontinuetorunlegacyapplicationsin thepre-
| Wildfire4.0    | mode,settheconfigurationoptionretrieve_instance_ |                            |     |     |     |     |
| -------------- | ------------------------------------------------ | -------------------------- | --- | --- | --- | --- |
| dependenciesto |                                                  | instance_and_generic_deps. |     |     |     |     |
ThepropertyIpfcModel.Type returnsthetypeofmodelin theform ofthe
| IpfcModelTypeobject.The        |     | typesof | modelsareasfollows: |     |     |     |
| ------------------------------ | --- | ------- | ------------------- | --- | --- | --- |
| • EpfcMDL_ASSEMBLY—Specifiesan |     |         | assembly.           |     |     |     |
EpfcMDL_PART—Specifiesapart.
•
| • EpfcMDL_DRAWING—Specifies |     | adrawing. |     |                                 |     |     |
| --------------------------- | --- | --------- | --- | ------------------------------- | --- | --- |
| 92                          |     |           |     | Creo®ParametricVBAPIUser’sGuide |     |     |

| •   | EpfcMDL_2D_SECTION—Specifiesa2D |     |     |     |     | section. |     |     |
| --- | ------------------------------- | --- | --- | --- | --- | -------- | --- | --- |
• EpfcMDL_LAYOUT—Specifiesanotebook.
| •   | EpfcMDL_DWG_FORMAT—Specifiesadrawing |     |     |                     |     | format. |     |     |
| --- | ------------------------------------ | --- | --- | ------------------- | --- | ------- | --- | --- |
| •   | EpfcMDL_MFG—Specifiesa               |     |     | manufacturingmodel. |     |         |     |     |
• EpfcMDL_REPORT—Specifiesareport.
| •   | EpfcMDL_MARKUP—Specifiesadrawing |     |     |     |          | markup. |     |     |
| --- | -------------------------------- | --- | --- | --- | -------- | ------- | --- | --- |
| •   | EpfcMDL_DIAGRAM—Specifies        |     |     |     | adiagram |         |     |     |
• EpfcMDL_CE_SOLID—SpecifiesaLayoutmodel.
Note
the VBAPI methodswillonlybe ableto readmodelsof typeLayout,but
|     | willnotbe               |                            | ableto passLayout |                    | modelsas     | inputto                | othermethods.PTC |        |
| --- | ----------------------- | -------------------------- | ----------------- | ------------------ | ------------ | ---------------------- | ---------------- | ------ |
|     | recommendsthatyoureview |                            |                   |                    | allthe VBAPI | applicationsthatusethe |                  |        |
|     | objectIpfcModelType     |                            |                   | andmodifythecodeas |              |                        | appropriateto    | ensure |
|     | thatthe                 | applicationsworkcorrectly. |                   |                    |              |                        |                  |        |
ThepropertyIpfcModel.IsModified identifieswhetherthemodelhasbeen
| modifiedsinceitwas |     |     | last saved. |     |     |     |     |     |
| ------------------ | --- | --- | ----------- | --- | --- | --- | --- | --- |
ThepropertyIpfcModel.Version returnsthe versionof thespecifiedmodel
| fromthePDM                    |     | system.It | canbe  | NULL,     | ifnotset.                  |         |         |     |
| ----------------------------- | --- | --------- | ------ | --------- | -------------------------- | ------- | ------- | --- |
| ThepropertyIpfcModel.Revision |     |           |        |           | returnstherevisionnumberof |         |         | the |
| specifiedmodelfrom            |     |           | thePDM | system.It | canbe                      | NULL,if | notset. |     |
ThepropertyIpfcModel.Branch
|            |     |           |       |       | returnsthebranch |     | ofthespecifiedmodel |     |
| ---------- | --- | --------- | ----- | ----- | ---------------- | --- | ------------------- | --- |
| fromthePDM |     | system.It | canbe | NULL, | ifnotset.        |     |                     |     |
ThepropertyIpfcModel.ReleaseLevel returnsthe releaselevel ofthe
| specifiedmodelfrom |     |     | thePDM | system.It | canbe | NULL,if | notset. |     |
| ------------------ | --- | --- | ------ | --------- | ----- | ------- | ------- | --- |
ThepropertyIpfcModel.VersionStamp returnsthe versionstampof the
specifiedmodel. Theversion stampis aCreo Parametricspecificidentifier that
| changeswith |     | eachchangemadeto |     |     | themodel. |     |     |     |
| ----------- | --- | ---------------- | --- | --- | --------- | --- | --- | --- |
ThemethodIpfcModel.ListDependencies()
|     |     |     |     |     |     |     | returns alistof | thefirst- |
| --- | --- | --- | --- | --- | --- | --- | --------------- | --------- |
leveldependencies forthe specifiedmodelin theCreo Parametricworkspacein
| theformof |     | theIpfcDependencies |     |     | object. |     |     |     |
| --------- | --- | ------------------- | --- | --- | ------- | --- | --- | --- |
Models 93

| Usethe methodIpfcModel.CleanupDependencies()to |            |         |                      |     | clean the |
| ---------------------------------------------- | ---------- | ------- | -------------------- | --- | --------- |
| dependenciesfor                                | anobjectin | theCreo | Parametricworkspace. |     |           |
Note
| Donot callthemethodIpfcModel.CleanupDependencies() |     |     |     |     | during |
| -------------------------------------------------- | --- | --- | --- | --- | ------ |
operationsthatalterthe dependencies,suchas,restructuring componentsand
creatingorredefiningfeatures.
ThemethodIpfcModel.ListDeclaredModels() returns alist ofallthe
| first-levelobjectsdeclaredforthe       |     | specifiedmodel. |     |              |        |
| -------------------------------------- | --- | --------------- | --- | ------------ | ------ |
| ThemethodIpfcModel.CheckIsModifiable() |     |                 |     | identifiesif | agiven |
modelcan bemodifiedwithout checkingfor anysubordinate models.This method
takesabooleanargumentShowUIthatdetermineswhethertheCreo Parametric
conflictresolutiondialogboxshouldbe displayedto resolveconflicts, if detected.
Ifthisargumentisfalse, thentheconflictresolutiondialogboxisnotdisplayed,
andthemodelcanbe modifiedonly ifthereareno conflictsthatcannotbe
overridden,or areresolvedby defaultresolutionactions.Fora genericmodel,if
| ShowUIistrue,thenall                    | instancesof |     | themodelarealsochecked. |            |           |
| --------------------------------------- | ----------- | --- | ----------------------- | ---------- | --------- |
| ThemethodIpfcModel.CheckIsSaveAllowed() |             |     |                         | identifies | ifa given |
modelcan besavedalongwith all ofitssubordinatemodels. The subordinate
modelscan besavedbasedon theirmodificationstatusandthe valueof the
save_objects.
| configurationoption |     |     | This methodalsochecks |     | thecurrentuser |
| ------------------- | --- | --- | --------------------- | --- | -------------- |
interfacecontextto identifyif itiscurrentlysafeto savethe model.Thus,calling
thismethodatdifferenttimes mightreturndifferent results.Thismethodtakesa
booleanargument ShowUI.Refertothe previousmethodformoreinformation on
thisargument.
| Model Operations |     |     |     |     |     |
| ---------------- | --- | --- | --- | --- | --- |
MethodsandPropertyIntroduced:
• IpfcModel.Backup()
• IpfcModel.Copy()
• IpfcModel.CopyAndRetrieve()
• IpfcModel.Rename()
• IpfcModel.Save()
• IpfcModel.Erase()
• IpfcModel.EraseWithDependencies()
• IpfcModel.Delete()
| 94  |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

• IpfcModel.Display()
• IpfcModel.CommonName
Thesemodeloperationsduplicatemostofthe commandsavailablein theCreo
ParametricFilemenu.
ThemethodIpfcModel.Backup() makesa backupof anobjectin memory to
adiskin aspecifieddirectory.
ThemethodIpfcModel.Copy() copiesthe specifiedmodelto anotherfile.
ThemethodIpfcModel.CopyAndRetrieve() copiesthe modelto another
name,andretrievesthatnewmodelintosession.
| ThemethodIpfcModel.Rename() |     | renamesaspecifiedmodel. |                  |     |        |     |
| --------------------------- | --- | ----------------------- | ---------------- | --- | ------ | --- |
| ThemethodIpfcModel.Save()   |     | storesthe               | specifiedmodelto |     | adisk. |     |
ThemethodIpfcModel.Erase() erasesthespecifiedmodelfrom thesession.
Modelsusedby othermodelscannotbe eraseduntil themodelsdependentupon
themareerased.
ThemethodIpfcModel.EraseWithDependencies() erasesthespecified
modelfrom thesessionandall themodelson whichthespecifiedmodeldepends
fromdisk,if thedependencies arenotneededbyotheritems in session.
Note
| However, whileerasing             | anactivemodel, |     | IpfcModel.Erase() |           |            | and |
| --------------------------------- | -------------- | --- | ----------------- | --------- | ---------- | --- |
| IpfcModel.EraseWithDependencies() |                |     |                   | onlyclear | thegraphic |     |
displayimmediately,theydo notclearthedatain thememory untilthe control
returnsto Creo Parametricfromthe theVB APIapplication. Therefore,after
calling themthecontrolmustbe returnedto Creo Parametricbeforecalling
| anyotherfunction,otherwise |     | thebehavior | ofCreo | Parametricmaybe |     |     |
| -------------------------- | --- | ----------- | ------ | --------------- | --- | --- |
unpredictable.
| ThemethodIpfcModel.Delete() |     | removesthespecifiedmodelfrom |     |     |     |     |
| --------------------------- | --- | ---------------------------- | --- | --- | --- | --- |
memoryanddisk.
ThemethodIpfcModel.Display() displaysthe specifiedmodel.Youmust
callthismethodif youcreateanewwindowfor amodelbecausethemodelwill
untilyoucallIpfcDisplay.
notbedisplayedin thewindow
| ThepropertyIpfcModel.CommonName |     |     | modifiesthe | commonname |     | ofthe |
| ------------------------------- | --- | --- | ----------- | ---------- | --- | ----- |
specifiedmodel. You canmodifythis namefor modelsthatare notyetownedby
WindchillPDMLink, orin certainsituationsif theconfigurationoption let_
| proe_rename_pdm_objectsissetto |     |     | yes. |     |     |     |
| ------------------------------ | --- | --- | ---- | --- | --- | --- |
| Models                         |     |     |      |     |     | 95  |

Running Creo ModelCHECK
CreoModelCHECKisan integratedapplicationthatrunstransparently within
CreoParametric.Creo ModelCHECKusesa configurablelist ofcompany design
standardsandbestmodeling practices.Youcan configureCreo ModelCHECKto
runinteractivelyorautomatically whenyouregenerateorsave amodel.
MethodsandPropertiesIntroduced:
• IpfcBaseSession.ExecuteModelCheck()
• CCpfcModelCheckInstructions.Create()
• IpfcModelCheckInstructions.ConfigDir
• IpfcModelCheckInstructions.Mode
• IpfcModelCheckInstructions.OutputDir
• IpfcModelCheckInstructions.ShowInBrowser
• IpfcModelCheckResults.NumberOfErrors
• IpfcModelCheckResults.NumberOfWarnings
• IpfcModelCheckResults.WasModelSaved
You canrunCreoModelCHECKfrom an externalapplicationusing themethod
IpfcBaseSession.ExecuteModelCheck().This methodtakesthemodel
Modelon whichyouwantto runCreoModelCHECKand instructionsin theform
oftheobjectIpfcModelCheckInstructions asitsinput parameters.This
objectcontains thefollowingparameters:
• ConfigDir—Specifiesthe locationoftheconfigurationfiles. Ifthis
parameteris settoNULL, thedefault CreoModelCHECKconfigurationfiles
areused.
• Mode—Specifies themodein whichyouwant torunCreo ModelCHECK.
Themodes are:
○ MODELCHECK_GRAPHICS—Interactivemode
○ MODELCHECK_NO_GRAPHICS—Batchmode
• OutputDir—Specifiesthe locationforthe reports.Ifyousetthisparameter
toNULL, thedefault CreoModelCHECKdirectory, asperconfig_
init.mc,willbe used.
• ShowInBrowser—Specifiesif theresults reportshouldbe displayedinthe
Web browser.
ThemethodCCpfcModelCheckInstructions.Create() createsthe
IpfcModelCheckInstructionsobjectcontainingthe CreoModelCHECK
instructionsdescribed above.
96 Creo®ParametricVBAPIUser’sGuide

Usethe methodsandproperties
IpfcModelCheckInstructions.ConfigDir,
IpfcModelCheckInstructions.Mode,
| IpfcModelCheckInstructions.OutputDir,    |     |     | and              |
| ---------------------------------------- | --- | --- | ---------------- |
| IpfcModelCheckInstructions.ShowInBrowser |     |     | to modifytheCreo |
ModelCHECKinstructions.
| ThemethodIpfcBaseSession.ExecuteModelCheck() |                     |               | returnsthe |
| -------------------------------------------- | ------------------- | ------------- | ---------- |
| resultsofthe                                 | CreoModelCHECKrunin | theform ofthe |            |
IpfcModelCheckResults object.This objectcontainsthe following
parameters:
| • NumberOfErrors—Specifiesthenumberof   |     | errorsdetected. |     |
| --------------------------------------- | --- | --------------- | --- |
| • NumberOfWarnings—Specifiesthenumberof |     | warningsfound.  |     |
• WasModelSaved—Specifieswhetherthe modelis savedwithupdates.
Usethe propertiesIpfcModelCheckResults.NumberOfErrors,
IpfcModelCheckResults.NumberOfWarning,and
IpfcModelCheckResults.WasModelSavedto accessthe resultsobtained.
| Custom | Checks |     |     |
| ------ | ------ | --- | --- |
Thissectiondescribeshowto define customchecksin CreoModelCHECKthat
userscanrunusingthe standardCreo ModelCHECKinterface inCreo Parametric.
| Todefine                     | andregisteracustomcheck: |                     |             |
| ---------------------------- | ------------------------ | ------------------- | ----------- |
| 1. SettheCUSTMTK_CHECKS_FILE |                          | configurationoption | in thestart |
configurationfileto atextfilethatstores thecheckdefinition. Forexample:
| CUSTMTK_CHECKS_FILE | text/custmtk_checks.txt. |     |     |
| ------------------- | ------------------------ | --- | --- |
2. Setthecontents ofthe CUSTMTK_CHECKS_FILE fileto define thechecks.
Eachcheckshouldlist thefollowingitems:
• DEF_<checkname>—Specifies thenameof thecheck.The formatmust
| be                          | CHKTK_<checkname>_<mode>,wheremodeis |                | PRT,ASM,orDRW. |
| --------------------------- | ------------------------------------ | -------------- | -------------- |
| • TAB_<checkname>—Specifies |                                      | thetabcategory | in theCreo     |
ModelCHECKreportunderwhichthecheck isclassified.Validtab values
are:
| ○   | INFO      |     |     |
| --- | --------- | --- | --- |
| ○   | PARAMETER |     |     |
| ○   | LAYER     |     |     |
| ○   | FEATURE   |     |     |
| ○   | RELATION  |     |     |
| ○   | DATUM     |     |     |
Models 97

| ○ MISC  |     |     |     |     |     |
| ------- | --- | --- | --- | --- | --- |
| ○ VDA   |     |     |     |     |     |
| ○ VIEWS |     |     |     |     |     |
• MSG_<checkname>—Specifies thedescriptionofthe checkthatappears
| in thelowerpartof |     | theCreo ModelCHECKreportwhenyou |     |     | selectthe |
| ----------------- | --- | ------------------------------- | --- | --- | --------- |
name.
• DSC_<checkname>—Specifies thenameof thecheck asitappears in
| theCreo                   | ModelCHECKreporttable. |     |                  |              |      |
| ------------------------- | ---------------------- | --- | ---------------- | ------------ | ---- |
| • ERM_<checkname>—Ifsetto |                        |     | INFO, thecheckis | consideredan | INFO |
checkand thereporttable displaysthetextfromthe firstitem returnedby
thecheck,insteadof acountof theitems.Otherwise,this valuemustbe
| included,    | butisignoredby | CreoParametric.         |     |              |         |
| ------------ | -------------- | ----------------------- | --- | ------------ | ------- |
| SeeExample1: |                | TextFileforCustomChecks |     | onpage102for | asample |
customcheckstextfile.
3. Addthecheckand itsvaluesto theCreo ModelCHECKconfigurationfile.
4. Registerthe CreoModelCHECKcheck fromthe VBAPI application.
Note
Otherthantherequirements listedabove,theVB API customchecks donot
haveaccessto therestof thevaluesin the CreoModelCHECKconfiguration
files. Allthecustomsettingsspecificto thecheck, suchasstartparameters,
constants,andso on,mustbe supportedby theuserapplicationandnotCreo
ModelCHECK.
| Registering | Custom | Checks |     |     |     |
| ----------- | ------ | ------ | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcBaseSession.RegisterCustomModelCheck
• CCpfcCustomCheckInstructions_Create
• IpfcCustomCheckInstructions.CheckName
• IpfcCustomCheckInstructions.CheckLabel
• IpfcCustomCheckInstructions.Listener
• IpfcCustomCheckInstructions.ActionButtonLabel
• IpfcCustomCheckInstructions.UpdateButtonLabel
| 98  |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

| ThemethodIpfcBaseSession.RegisterCustomModelCheck |     |     |     | registers |
| ------------------------------------------------- | --- | --- | --- | --------- |
acustomcheckthatcanbe includedin anyModelCHECKrun.This methodtakes
theinstructionsinthe formof theIpfcCustomCheckInstructionsobject
| asitsinputargument.              | Thisobjectcontains |           | thefollowingparameters: |     |
| -------------------------------- | ------------------ | --------- | ----------------------- | --- |
| • CheckName—Specifiesthe         |                    | nameofthe | customcheck.            |     |
| • CheckLabel—Specifiesthelabelof |                    |           | thecustomcheck.         |     |
• Listener—Specifiesthelistenerobjectcontainingthe customcheck methods.
| Refertothe | sectionCustomCheckListeners |     | onpage100for | more |
| ---------- | --------------------------- | --- | ------------ | ---- |
information.
• ActionButtonLabel—Specifiesthe labelfor theactionbutton. Ifyouspecify
| NULLforthis | parameter,this | buttonis | notshown. |     |
| ----------- | -------------- | -------- | --------- | --- |
• UpdateButtonLabel—Specifiesthelabel forthe updatebutton. Ifyouspecify
| NULLforthis                                  | parameter,this | buttonis | notshown. |            |
| -------------------------------------------- | -------------- | -------- | --------- | ---------- |
| ThemethodCCpfcCustomCheckInstructions_Create |                |          |           | createsthe |
IpfcCustomCheckInstructionsobjectcontainingthecustomcheck
| instructionsdescribed | above. |     |     |     |
| --------------------- | ------ | --- | --- | --- |
Usethe propertiesIpfcCustomCheckInstructions.SetCheckName,
IpfcCustomCheckInstructions.SetCheckLabel,
IpfcCustomCheckInstructions.SetListener,
| IpfcCustomCheckInstructions.SetActionButtonLabel, |     |     |     | and |
| ------------------------------------------------- | --- | --- | --- | --- |
IpfcCustomCheckInstructions.SetUpdateButtonLabelto modify
theinstructions.
Thefollowingfigureillustrateshowtheresultsof somecustomchecks mightbe
| displayedin | theModelCHECKreport. |     |     |     |
| ----------- | -------------------- | --- | --- | --- |
Models 99

Custom Check Listeners
MethodsandPropertiesIntroduced:
• IpfcModelCheckCustomCheckListener.OnCustomCheck
• CCpfcCustomCheckResults_Create
• IpfcCustomCheckResults.ResultsCount
• IpfcCustomCheckResults.ResultsTable
• IpfcCustomCheckResults.ResultsUrl
• IpfcModelCheckCustomCheckListener.OnCustomCheckAction
• IpfcModelCheckCustomCheckListener.OnCustomCheckUpdate
TheinterfaceIpfcModelCheckCustomCheckListener providesthe
methodsignaturestoimplementacustomCreo ModelCHECKcheck.
Eachlistenermethodtakesthe followinginput arguments:
• CheckName—Thenameof thecustomcheckas definedin theoriginal callto
themethodIpfcBaseSession.RegisterCustomModelCheck
• Mdl—Themodelbeingchecked.
100 Creo®ParametricVBAPIUser’sGuide

Theapplicationmethodthatoverrides
| IpfcModelCheckCustomCheckListener.OnCustomCheck |     |     |     | isusedto |
| ----------------------------------------------- | --- | --- | --- | -------- |
evaluateacustomdefinedcheck.The userapplicationrunsthecheck onthe
| specifiedmodelandreturns | theresultsin | theform | ofthe |     |
| ------------------------ | ------------ | ------- | ----- | --- |
IpfcCustomCheckResultsobject
• ResultsCount—Specifiesan integerindicatingthenumberof errorsfound by
thecheck.This valueis displayedin the CreoModelCHECKreportgenerated.
• ResultsTable—Specifiesa listof textdescriptionsof theproblemencountered
| foreacherror | orwarning. |     |     |     |
| ------------ | ---------- | --- | --- | --- |
• ResultsUrl—Specifiesthe linktoan application-ownedpagethatprovides
| informationon                                        | theresultsof | thecustomcheck. |            |         |
| ---------------------------------------------------- | ------------ | --------------- | ---------- | ------- |
| ThemethodIpfcIpfcCustomCheckResults_Create           |              |                 | createsthe |         |
| IpfcCustomCheckResultsobjectcontainingthecustomcheck |              |                 |            | results |
describedabove.
Usethe propertiesIpfcCustomCheckResults.SetResultsCount,
| IpfcCustomCheckResults.SetResultsTable, |                  |     | and         |             |
| --------------------------------------- | ---------------- | --- | ----------- | ----------- |
| IpfcCustomCheckResults.SetResultsUrl    |                  |     | listedabove | tomodifythe |
| customchecks                            | resultsobtained. |     |             |             |
Themethodthatoverrides
IpfcModelCheckCustomCheckListener.OnCustomCheckAction
is
calledwhenthecustomcheck’s Actionbutton ispressed.Theinput supplied
| includesthetextselectedby |     | theuserfrom thecustomcheckresults. |     |     |
| ------------------------- | --- | ---------------------------------- | --- | --- |
Themethodthatoverrides
IpfcModelCheckCustomCheckListener.OnCustomCheckUpdate is
calledwhenthecustomcheck’s Updatebutton ispressed.Theinput supplied
| includesthetextselectedby |     | theuserfrom thecustomcheckresults. |     |     |
| ------------------------- | --- | ---------------------------------- | --- | --- |
CustomCreo ModelCHECKcheckscan haveanActionbutton to highlightthe
| problem,andpossibly | anUpdatebutton | to fixitautomatically. |     |     |
| ------------------- | -------------- | ---------------------- | --- | --- |
ThefollowingfiguredisplaystheCreo ModelCHECKreportwith an Action
| buttonthatinvokes | the |     |     |     |
| ----------------- | --- | --- | --- | --- |
IpfcModelCheckCustomCheckListener.OnCustomCheckAction
Models 101

| Example                                | 1: Text | File                       | for Custom | Checks |                  |     |
| -------------------------------------- | ------- | -------------------------- | ---------- | ------ | ---------------- | --- |
| Thefollowingisthe                      |         | textfilecustmtk_checks.txt |            |        | locatedat<creo_  |     |
| vbapi_loadpoint>/vbapi_examples/models |         |                            |            |        | for customchecks |     |
examples.
| # Custom   | TK Check | File  |               |     |     |     |
| ---------- | -------- | ----- | ------------- | --- | --- | --- |
| # def-name | of       | check | as registered |     |     |     |
# MDLPARAM_NAME
| DEF_MDLPARAM_NAME |     | CHKTK_MDLPARAM_NAME_PRT |         |           |            |        |
| ----------------- | --- | ----------------------- | ------- | --------- | ---------- | ------ |
| TAB_MDLPARAM_NAME |     | DATUM                   |         |           |            |        |
| MSG_MDLPARAM_NAME |     | CUSTOM                  | : Datum | - Model   | Param Name |        |
| ERM_MDLPARAM_NAME |     | CUSTOM                  | : Datum | - Invalid | Parameter  | value. |
DSC_MDLPARAM_NAME CUSTOM : Datum - Check Model Parameter Name.
| 102 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

# MODEL_ACCURACY
| DEF_MODEL_ACCURACY |     | CHKTK_MODEL_ACCURACY_PRT |     |     |     |     |
| ------------------ | --- | ------------------------ | --- | --- | --- | --- |
| TAB_MODEL_ACCURACY |     | DATUM                    |     |     |     |     |
MSG_MODEL_ACCURACY CUSTOM : Datum - Report Model accuracy type
ERM_MODEL_ACCURACY CUSTOM : Datum - Report Model accuracy type
DSC_MODEL_ACCURACY CUSTOM : Datum - Report Model accuracy type
# DWGVIEW_GENERIC
| DEF_DWGVIEW_GENERIC |     | CHKTK_DWGVIEW_GENERIC_DRW |     |     |     |     |
| ------------------- | --- | ------------------------- | --- | --- | --- | --- |
| TAB_DWGVIEW_GENERIC |     | VIEWS                     |     |     |     |     |
MSG_DWGVIEW_GENERIC CUSTOM : Drawing views containing generic
models:
| ERM_DWGVIEW_GENERIC                               |                                        | N/A    |           |       |                |        |
| ------------------------------------------------- | -------------------------------------- | ------ | --------- | ----- | -------------- | ------ |
| DSC_DWGVIEW_GENERIC                               |                                        | CUSTOM | : Drawing | Views | using Generics |        |
| Example                                           | 2: Registering                         |        | Custom    | Creo  | ModelCHECK     | Checks |
| Thesamplecodein                                   | thefilepfcModelCheckExamples.vblocated |        |           |       |                | at     |
| <creo_vbapi_loadpoint>/vbapi_examplesdemonstrates |                                        |        |           |       |                | howto  |
registercustomCreo ModelCHECKchecksusing theVB API.The following
| customchecks | areregistered: |     |     |     |     |     |
| ------------ | -------------- | --- | --- | --- | --- | --- |
• CHKTK_MDLPARAM_NAME—Determines ifthe modelhasa parameterwhose
| nameisequal | to themodelname. |     |     |     |     |     |
| ----------- | ---------------- | --- | --- | --- | --- | --- |
• CHKTK_MODEL_ACCURACY—Checksthetypeof accuracydefinedfor the
model.
| • CHKTK_DWGVIEW_GENERIC—Drawing              |                                        |     |         | modecheckthatidentifies |               | the   |
| -------------------------------------------- | -------------------------------------- | --- | ------- | ----------------------- | ------------- | ----- |
| drawingviews                                 | thatusegenericmodels.                  |     |         |                         |               |       |
| Example                                      | 3: Implementing                        |     | a Model | Name                    | Parameter     | Check |
| Thesamplecodein                              | thefilepfcModelCheckExamples.vblocated |     |         |                         |               | at    |
| <creo_vbapi_loadpoint>/vbapi_examplesdefines |                                        |     |         |                         | thecustomCreo |       |
ModelCHECKcheck forthe parametername in amodel.This checkupdatesthe
parameternameto be equalto themodelname,if theparameter existswitha
differentname,orcreatesthe parameterwith themodelnameif itdoesnotexist.
| Example                                      | 4: Implementing                        |     | a Model | Accuracy | Type          | Check |
| -------------------------------------------- | -------------------------------------- | --- | ------- | -------- | ------------- | ----- |
| Thesamplecodein                              | thefilepfcModelCheckExamples.vblocated |     |         |          |               | at    |
| <creo_vbapi_loadpoint>/vbapi_examplesdefines |                                        |     |         |          | thecustomCreo |       |
ModelCHECKcheck forthe typeofaccuracywhetherrelativeorabsolutethathas
beensetfor amodel. Thischeck hasachecklistener method,butnoactionor
| updatelistener | methodsinceitisan |     | info-only | check. |     |     |
| -------------- | ----------------- | --- | --------- | ------ | --- | --- |
Models 103

| Example                                      | 5:     | Implementing                           |     | a Check |     | for Drawing | Views         | Using |
| -------------------------------------------- | ------ | -------------------------------------- | --- | ------- | --- | ----------- | ------------- | ----- |
| Generic                                      | Models |                                        |     |         |     |             |               |       |
| Thesamplecodein                              |        | thefilepfcModelCheckExamples.vblocated |     |         |     |             |               | at    |
| <creo_vbapi_loadpoint>/vbapi_examplesdefines |        |                                        |     |         |     |             | thecustomCreo |       |
ModelCHECKcheck foridentifying drawingviews using genericmodels.This
checkhasachecklistener methodandan actionlistenermethod tohighlight the
viewsthatusegenericmodels.
| Example       | 6:  | Changes | to        | the Creo | ModelCHECK |        |     |     |
| ------------- | --- | ------- | --------- | -------- | ---------- | ------ | --- | --- |
| Configuration |     | Files   | to enable | Custom   |            | Checks |     |     |
Lines added to the ModelCheck configuration file (default_checks.
mch)
| E                       | Check  | the item. | If        | not succeed, |     | report | as an error  |             |
| ----------------------- | ------ | --------- | --------- | ------------ | --- | ------ | ------------ | ----------- |
| W                       | Check  | the item. | If        | not succeed, |     | report | as a warning |             |
| N                       | Do not | check     | the item. |              |     |        |              |             |
| Y                       | Check  | the item. | If        | not succeed, |     | do not | report       | err or warn |
| CHKTK_MDLPARAM_NAME_PRT |        |           |           | YNEW         |     | E      | E            | E E         |
Y
| CHKTK_MODEL_ACCURACY_PRT |     |     |     | YNEW |     | Y   | Y   | Y Y |
| ------------------------ | --- | --- | --- | ---- | --- | --- | --- | --- |
Y
| CHKTK_DWGVIEW_GENERIC_DRW |     |     |     | YNEW |     | E   | E   | E E |
| ------------------------- | --- | --- | --- | ---- | --- | --- | --- | --- |
Y
| Lines added         | to  | the ModelCheck |                    | start | file | (sample_start.mcs)              |     |     |
| ------------------- | --- | -------------- | ------------------ | ----- | ---- | ------------------------------- | --- | --- |
| CUSTMTK_CHECKS_FILE |     |                | custmtk_checks.txt |       |      |                                 |     |     |
| 104                 |     |                |                    |       |      | Creo®ParametricVBAPIUser’sGuide |     |     |

10
Drawings
OverviewofDrawingsintheVBAPI.........................................................................106
CreatingDrawingsfromTemplates...........................................................................106
ObtainingDrawingModels.......................................................................................108
DrawingInformation................................................................................................108
DrawingOperations.................................................................................................109
DrawingSheets.......................................................................................................110
DrawingViews........................................................................................................113
DrawingDimensions................................................................................................118
DrawingTables.......................................................................................................124
DetailItems.............................................................................................................132
DetailEntities..........................................................................................................134
OLEObjects...........................................................................................................136
DetailNotes............................................................................................................136
DetailGroups..........................................................................................................140
DetailSymbols........................................................................................................141
DetailAttachments..................................................................................................151
Thischapterdescribes howtoprogramdrawing functions usingthe VBAPI.
105

| Overview | of Drawings | in the | VB  | API |     |
| -------- | ----------- | ------ | --- | --- | --- |
Thissectiondescribesthe functionsthatdealwith drawings.Youcan create
drawingsof allCreo Parametricmodelsusingthe functionsin theVB API.You
canannotatethedrawing, manipulatedimensions,and uselayers tomanage the
| displayof differentitems. |     |     |     |     |     |
| ------------------------- | --- | --- | --- | --- | --- |
Unlessotherwise specified,the VBAPI functionsthatoperate ondrawings use
worldunits.
| Creating | Drawings | from Templates |     |     |     |
| -------- | -------- | -------------- | --- | --- | --- |
Drawingtemplatessimplify theprocessofcreatinga drawingusing theVB API.
CreoParametriccan createviews, settheview display,createsnaplines, and
| showthe modeldimensionsbasedonthe |     | template.Use | templatesto: |     |     |
| --------------------------------- | --- | ------------ | ------------ | --- | --- |
• Definelayoutviews
| • Setview display |     |     |     |     |     |
| ----------------- | --- | --- | --- | --- | --- |
• Placenotes
• Placesymbols
• Definetables
• Showdimensions
MethodIntroduced:
• IpfcBaseSession.CreateDrawingFromTemplate()
Usethe methodIpfcBaseSession.CreateDrawingFromTemplate()
tocreateadrawing from thedrawing templateandto returnthecreateddrawing.
Theattributesare:
• Newdrawingname
| • Nameofan | existing template |     |     |     |     |
| ---------- | ----------------- | --- | --- | --- | --- |
• Nameandtypeofthe solidmodelto usewhilepopulatingtemplateviews
• Sequence ofoptions tocreatethe drawing.The optionsareas follows:
| ○ EpfcDRAWINGCREATE_DISPLAY_DRAWING—display |     |     |     | thenew |     |
| ------------------------------------------- | --- | --- | --- | ------ | --- |
drawing.
| ○ EpfcDRAWINGCREATE_SHOW_ERROR_DIALOG—display |     |     |     | theerror |     |
| --------------------------------------------- | --- | --- | --- | -------- | --- |
dialogbox.
| ○ EpfcDRAWINGCREATE_WRITE_ERROR_FILE—write |     |     |     | the errorsto | a   |
| ------------------------------------------ | --- | --- | --- | ------------ | --- |
file.
| ○ EpfcDRAWINGCREATE_PROMPT_UNKNOWN_PARAMS—prompt |                     |            |                                 |     | the |
| ------------------------------------------------ | ------------------- | ---------- | ------------------------------- | --- | --- |
| useron                                           | encounteringunknown | parameters |                                 |     |     |
| 106                                              |                     |            | Creo®ParametricVBAPIUser’sGuide |     |     |

| Drawing |     | Creation | Errors |     |     |     |
| ------- | --- | -------- | ------ | --- | --- | --- |
Theexception IpfcXToolkitDrawingCreateErrorsisthrown ifan error
isencounteredwhencreatinga drawingfrom atemplate. Thisexception contains
| alistof | errorswhichoccurred |     | duringdrawing | creation. |     |     |
| ------- | ------------------- | --- | ------------- | --------- | --- | --- |
Note
Whenthis exceptiontypeisencountered,thedrawing isactually created,but
|     | someofthe | contentsfailedto | generatecorrectly. |     |     |     |
| --- | --------- | ---------------- | ------------------ | --- | --- | --- |
Theexception messagewilllist thedetailsfor eacherrorincluding itstype,sheet
number, viewname,and(if applicable) itemname.The typesoferrorsare as
follows:
| •   | Type—Thetypeoferroras |     | follows: |     |     |     |
| --- | --------------------- | --- | -------- | --- | --- | --- |
○ EpfcDWGCREATE_ERR_SAVED_VIEW_DOESNT_EXIST—Saved
|     | view                                             | doesnotexist. |     |     |     |       |
| --- | ------------------------------------------------ | ------------- | --- | --- | --- | ----- |
|     | ○ EpfcDWGCREATE_ERR_X_SEC_DOESNT_EXIST—Specified |               |     |     |     | cross |
sectiondoesnotexist.
|     | ○ EpfcDWGCREATE_ERR_EXPLODE_DOESNT_EXIST—Exploded |     |     |     |     | state |
| --- | ------------------------------------------------- | --- | --- | --- | --- | ----- |
didnotexist.
|     | ○ EpfcDWGCREATE_ERR_MODEL_NOT_EXPLODABLE—Model |     |               |     |             | cannot |
| --- | ---------------------------------------------- | --- | ------------- | --- | ----------- | ------ |
|     | be exploded.                                   |     |               |     |             |        |
|     | ○ EpfcDWGCREATE_ERR_SEC_NOT_PERP—Cross         |     |               |     | sectionview | not    |
|     | perpendicularto                                |     | thegivenview. |     |             |        |
|     | ○ EpfcDWGCREATE_ERR_NO_RPT_REGIONS—Repeat      |     |               |     | regionsnot  |        |
available.
|     | ○ EpfcDWGCREATE_ERR_FIRST_REGION_USED—Repeat |        |                  |     |     | region was |
| --- | -------------------------------------------- | ------ | ---------------- | --- | --- | ---------- |
|     | unableto                                     | usethe | regionspecified. |     |     |            |
EpfcDWGCREATE_ERR_NOT_PROCESS_ASSEM—
|     | ○                                   |     |       |     | Modelisnot  | a      |
| --- | ----------------------------------- | --- | ----- | --- | ----------- | ------ |
|     | processassembly                     |     | view. |     |             |        |
|     | ○ EpfcDWGCREATE_ERR_NO_STEP_NUM—The |     |       |     | processstep | number |
doesnotexist.
|     | ○ EpfcDWGCREATE_ERR_TEMPLATE_USED—The |     |     |     | templatedoesnot |     |
| --- | ------------------------------------- | --- | --- | --- | --------------- | --- |
exist.
|     | ○ EpfcDWGCREATE_ERR_NO_PARENT_VIEW_FOR_PROJ—There |           |                    |       |     | isno   |
| --- | ------------------------------------------------- | --------- | ------------------ | ----- | --- | ------ |
|     | possibleparent                                    | viewfor   | thisprojectedview. |       |     |        |
|     | ○ EpfcDWGCREATE_ERR_CANT_GET_PROJ_PARENT—Could    |           |                    |       |     | notget |
|     | theprojected                                      | parentfor | adrawing           | view. |     |        |
Drawings 107

|     | ○ EpfcDWGCREATE_ERR_SEC_NOT_PARALLEL—The |               |     |                 |     | designated | cross |
| --- | ---------------------------------------- | ------------- | --- | --------------- | --- | ---------- | ----- |
|     | sectionwas                               | notparallelto |     | thecreatedview. |     |            |       |
○ EpfcDWGCREATE_ERR_SIMP_REP_DOESNT_EXIST—The
|          | designatedsimplifiedrepresentationdoes |                              |          |      | notexist.       |     |     |
| -------- | -------------------------------------- | ---------------------------- | -------- | ---- | --------------- | --- | --- |
| Example: | Drawing                                |                              | Creation | from | a Template      |     |     |
|          |                                        | thefilepfcDrawingExamples.vb |          |      | locatedat<creo_ |     |     |
Thesamplecodein
| vbapi_loadpoint>/vbapi_examplescreatesanewdrawing |     |     |     |     |     | using | a   |
| ------------------------------------------------- | --- | --- | --- | --- | --- | ----- | --- |
predefinedtemplate.
| Obtaining                 |     | Drawing |     | Models        |         |     |     |
| ------------------------- | --- | ------- | --- | ------------- | ------- | --- | --- |
| Thissectiondescribeshowto |     |         |     | obtaindrawing | models. |     |     |
MethodsIntroduced:
• IpfcBaseSession.RetrieveModel()
• IpfcBaseSession.GetModel()
• IpfcBaseSession.GetModelFromDescr()
• IpfcBaseSession.ListModels()
• IpfcBaseSession.ListModelsByType()
ThemethodIpfcBaseSession.RetrieveModel() retrieves thedrawing
specifiedby themodeldescriptor. Modeldescriptorsaredataobjectsusedto
describeamodelfileanditslocationinthe system.The methodreturnsthe
retrieveddrawing.
ThemethodIpfcBaseSession.GetModel() returnsa drawing basedon its
| nameandtype, |     | whereasIpfcBaseSession.GetModelFromDescr() |     |     |     |     |     |
| ------------ | --- | ------------------------------------------ | --- | --- | --- | --- | --- |
returnsadrawing specifiedbythe modeldescriptor.The modelmustbe in
session.
Usethe methodIpfcBaseSession.ListModels() toreturn asequenceof
| allthedrawings |     | in session. |     |     |     |     |     |
| -------------- | --- | ----------- | --- | --- | --- | --- | --- |
| Drawing        |     | Information |     |     |     |     |     |
MethodsandPropertyIntroduced:
• IpfcModel2D.ListModels()
• IpfcModel2D.GetCurrentSolid()
• IpfcModel2D.ListSimplifiedReps()
• IpfcModel2D.TextHeight
| 108 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- | --- |

ThemethodIpfcModel2D.ListModels() returnsa listof allthesolid
modelsused inthe drawing.
ThemethodIpfcModel2D.GetCurrentSolid() returnsthecurrentsolid
modelof thedrawing.
ThemethodIpfcModel2D.ListSimplifiedReps() returns thesimplified
representationsofasolidmodelthatareassignedto thedrawing.
ThepropertyIpfcModel2D.TextHeight returnsthe textheightof the
drawing.
Drawing Operations
MethodsIntroduced:
• IpfcModel2D.AddModel()
• IpfcModel2D.DeleteModel()
• IpfcModel2D.ReplaceModel()
• IpfcModel2D.SetCurrentSolid()
• IpfcModel2D.AddSimplifiedRep()
• IpfcModel2D.DeleteSimplifiedRep()
• IpfcModel2D.Regenerate()
• IpfcModel2D.CreateDrawingDimension()
• IpfcModel2D.CreateView()
ThemethodIpfcModel2D.AddModel() addsa newsolidmodelto the
drawing.
ThemethodIpfcModel2D.DeleteModel() removesamodelfromthe
drawing.The modelto be deletedshould notappearin anyof thedrawing views.
ThemethodIpfcModel2D.ReplaceModel() replaces amodelinthe
drawingwith arelated model(the relationshipshouldbe by familytableor
interchangeassembly).It allowsyouto replacemodelsthatareshownin drawing
viewsandregeneratestheview.
ThemethodIpfcModel2D.SetCurrentSolid() assignsthecurrentsolid
modelfor thedrawing. Beforecalling thismethod,thesolidmodelmustbe
assignedto the drawingusing themethod IpfcModel2D.AddModel().To
seethe changestoparametersandfields reflectingthechangeof thecurrentsolid
model,regeneratethedrawing usingthe method
IpfcSheetOwner.RegenerateSheet().
ThemethodIpfcModel2D.AddSimplifiedRep() associatesthe drawing
withthesimplifiedrepresentationof an assembly
Drawings 109

| ThemethodIpfcModel2D.DeleteSimplifiedRep() |     |     |     | removesthe |     |
| ------------------------------------------ | --- | --- | --- | ---------- | --- |
associationof thedrawing withan assembly simplifiedrepresentation.The
simplifiedrepresentation tobe deleted shouldnotappearin anyof thedrawing
views.
Usethe propertyIpfcModel2D.Regenerate()to regenerate thedrawing
draftentitiesandappearance.
| ThemethodIpfcModel2D.CreateDrawingDimension() |     |     |     |     | createsanew |
| --------------------------------------------- | --- | --- | --- | --- | ----------- |
drawingdimension basedon thedataobjectthatcontains informationaboutthe
locationof thedimension.This methodreturnsthe createddimension.Referto the
| sectionDrawingDimensionson        |     | page118. |                    |     |      |
| --------------------------------- | --- | -------- | ------------------ | --- | ---- |
| ThemethodIpfcModel2D.CreateView() |     |          | createsanewdrawing |     | view |
basedon thedataobjectthatcontains informationabouthowto createtheview.
Themethodreturns thecreateddrawing view.Referto thesectionCreating
| DrawingViews | onpage113. |         |       |            |     |
| ------------ | ---------- | ------- | ----- | ---------- | --- |
| Example:     | Replace    | Drawing | Model | Solid with | its |
Generic
| Thesamplecodein | thefilepfcDrawingExamples.vb |     |     | locatedat<creo_ |     |
| --------------- | ---------------------------- | --- | --- | --------------- | --- |
vbapi_loadpoint>/vbapi_examplesreplaces allsolidmodelinstances
ina drawingwith itsgeneric.Modelsarenotreplacedif thegenericmodelis
| alreadypresentin | thedrawing. |     |     |     |     |
| ---------------- | ----------- | --- | --- | --- | --- |
| Drawing          | Sheets      |     |     |     |     |
Adrawingsheetis representedby itsnumber.Drawing sheetsin theVB APIare
| identifiedby | thesamesheetnumbersseen |     | bya CreoParametricuser. |     |     |
| ------------ | ----------------------- | --- | ----------------------- | --- | --- |
Note
Theseidentifiersmaychangeif thesheetsare movedas aconsequenceof
| adding,removingor | reorderingsheets. |     |     |     |     |
| ----------------- | ----------------- | --- | --- | --- | --- |
| Drawing           | Sheet Information |     |     |     |     |
MethodsandPropertiesIntroduced
• IpfcSheetOwner.GetSheetTransform()
• IpfcSheetOwner.GetSheetInfo()
• IpfcSheetOwner.GetSheetScale()
| 110 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

• IpfcSheetOwner.GetSheetFormat()
• IpfcSheetOwner.GetSheetFormatDescr()
• IpfcSheetOwner.GetSheetBackgroundView()
• IpfcSheetOwner.NumberOfSheets
• IpfcSheetOwner.CurrentSheetNumber
• IpfcSheetOwner.GetSheetUnits()
SupersededMethod:
• IpfcSheetOwner.GetSheetData()
| ThemethodIpfcSheetOwner.GetSheetTransform() | returnsthe |     |
| ------------------------------------------- | ---------- | --- |
transformationmatrixforthe sheetspecifiedbythe sheetnumber.This
transformationmatrixincludesthe scalingneededtoconvertscreen coordinatesto
| drawingcoordinates(whichusethedesignated drawing | units).          |     |
| ------------------------------------------------ | ---------------- | --- |
| ThemethodIpfcSheetOwner.GetSheetInfo()           | returnssheetdata |     |
includingthesize,orientation,and unitsof thesheetspecifiedby thesheet
number.
ThemethodIpfcSheetOwner.GetSheetData()
andtheinterface
IpfcSheetDatahavebeendeprecated.Use themethod
| IpfcSheetOwner.GetSheetInfo()andthe interfaceIpfcSheetInfo |     |     |
| ---------------------------------------------------------- | --- | --- |
instead.
ThemethodIpfcSheetOwner.GetSheetScale() returns thescaleof the
drawingon aparticular sheetbasedonthe drawingmodelusedto measurethe
scale.If nomodelsareusedin thedrawing thenthedefault scalevalueis1.0.
| ThemethodIpfcSheetOwner.GetSheetFormat() | returnsthe | drawing |
| ---------------------------------------- | ---------- | ------- |
formatused forthe sheetspecifiedby thesheetnumber. Itreturns anullvalueif
noformatisassignedto thesheet.
| ThemethodIpfcSheetOwner.GetSheetFormatDescr() |     | returnsthe |
| --------------------------------------------- | --- | ---------- |
modeldescriptorof thedrawing formatused forthe specifieddrawingsheet.
| ThemethodIpfcSheetOwner.GetSheetBackgroundView() |     | returns the |
| ------------------------------------------------ | --- | ----------- |
viewobjectrepresentingthebackground viewof thesheetspecifiedbythe sheet
number.
| ThepropertyIpfcSheetOwner.NumberOfSheets | returnsthenumberof |     |
| ---------------------------------------- | ------------------ | --- |
sheetsinthe model.
ThepropertyIpfcSheetOwner.CurrentSheetNumber returnsthe current
sheetnumberin themodel.
Drawings 111

Note
| The sheetnumbersrangefrom1to |     | n,wherenisthe | numberofsheets. |     |
| ---------------------------- | --- | ------------- | --------------- | --- |
ThemethodIpfcSheetOwner.GetSheetUnits() returns theunitsusedby
thesheetspecifiedbythe sheetnumber.
| Drawing | Sheet Operations |     |     |     |
| ------- | ---------------- | --- | --- | --- |
MethodsIntroduced:
• IpfcSheetOwner.AddSheet()
• IpfcSheetOwner.DeleteSheet()
• IpfcSheetOwner.ReorderSheet()
• IpfcSheetOwner.RegenerateSheet()
• IpfcSheetOwner.SetSheetScale()
• IpfcSheetOwner.SetSheetFormat()
| ThemethodIpfcSheetOwner.AddSheet() |     |     | addsanewsheetto | themodel |
| ---------------------------------- | --- | --- | --------------- | -------- |
andreturnsthenumberof thenewsheet.
| ThemethodIpfcSheetOwner.DeleteSheet() |                    |           | removesthe | sheet |
| ------------------------------------- | ------------------ | --------- | ---------- | ----- |
| specifiedby                           | thesheetnumberfrom | themodel. |            |       |
Usethe methodIpfcSheetOwner.ReorderSheet()to reorderthe sheet
| fromaspecifiedsheetnumberto |     | anewsheetnumber. |     |     |
| --------------------------- | --- | ---------------- | --- | --- |
Note
The sheetnumberof otheraffectedsheetsalso changesduetoreorderingor
deletion.
ThemethodIpfcSheetOwner.RegenerateSheet() regeneratesthesheet
| specifiedby | thesheetnumber. |     |     |     |
| ----------- | --------------- | --- | --- | --- |
Note
| Youcanregenerateasheetonly |     | if itis displayed. |                                 |     |
| -------------------------- | --- | ------------------ | ------------------------------- | --- |
| 112                        |     |                    | Creo®ParametricVBAPIUser’sGuide |     |

Usethe methodIpfcSheetOwner.SetSheetScale() to setthescaleofa
modelon thesheetbasedon thedrawing modeltoscaleand thescaleto be used.
Passthe valueofthe DrawingModelparameter asnullto selectthe current
drawingmodel.
| Usethe methodIpfcSheetOwner.SetSheetFormat()to |     |     |     | applythe |
| ---------------------------------------------- | --- | --- | --- | -------- |
specifiedformatto adrawing sheetbasedonthe drawingformat, sheetnumberof
| theformat,and | thedrawing | model. |     |     |
| ------------- | ---------- | ------ | --- | --- |
Thesheetnumberofthe formatis specifiedbythe FormatSheetNumber
parameter. Thisnumberrangesfrom 1to thenumberof sheetsinthe format.Pass
| thevalueof | thisparameterasnulltousethe |     | firstformatsheet. |     |
| ---------- | --------------------------- | --- | ----------------- | --- |
Thedrawing modelisspecifiedbythe DrawingModelparameter.Pass thevalue
| ofthisparameter | asnullto                     | selectthecurrentdrawingmodel. |     |                 |
| --------------- | ---------------------------- | ----------------------------- | --- | --------------- |
| Example:        | Listing Drawing              | Sheets                        |     |                 |
| Thesamplecodein | thefilepfcDrawingExamples.vb |                               |     | locatedat<creo_ |
vbapi_loadpoint>/vbapi_examplesshowshowto listthe sheetsinthe
currentdrawing. Theinformation isplacedin anexternalbrowserwindow.
| Drawing | Views |     |     |     |
| ------- | ----- | --- | --- | --- |
Adrawingview isrepresented bythe interfaceIpfcView2D. Allmodelviewsin
thedrawing areassociative,thatis,if youchangeadimensional valuein oneview,
thesystemupdatesotherdrawing views accordingly.The modelautomatically
reflectsanydimensionalchangesthatyoumake toa drawing.In addition,
correspondingdrawingsalso reflectanychangesthatyoumaketo amodelsuchas
| theadditionordeletionoffeaturesanddimensional |         |       | changes. |     |
| --------------------------------------------- | ------- | ----- | -------- | --- |
| Creating                                      | Drawing | Views |          |     |
MethodIntroduced:
• IpfcModel2D.CreateView()
| ThemethodIpfcModel2D.CreateView() |     |     | createsanewview | in the |
| --------------------------------- | --- | --- | --------------- | ------ |
drawing.Beforecalling thismethod,the drawingmustbe displayedin awindow.
TheinterfaceIpfcView2DCreateInstructions contains detailsonhowto
createthe view. Thetypesof drawingviews supported forcreationare:
| • EpfcDRAWVIEW_GENERAL—General             |     |     | drawingviews |       |
| ------------------------------------------ | --- | --- | ------------ | ----- |
| • EpfcDRAWVIEW_PROJECTION—Projecteddrawing |     |     |              | views |
Drawings 113

| General | Drawing | Views |     |     |     |
| ------- | ------- | ----- | --- | --- | --- |
TheinterfaceIpfcGeneralViewCreateInstructions containsdetails on
| howto | creategeneraldrawing |     | views. |     |     |
| ----- | -------------------- | --- | ------ | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcGeneralViewCreateInstructions.Create()
• IpfcGeneralViewCreateInstructions.ViewModel
• IpfcGeneralViewCreateInstructions.Location
• IpfcGeneralViewCreateInstructions.SheetNumber
• IpfcGeneralViewCreateInstructions.Orientation
• IpfcGeneralViewCreateInstructions.Exploded
• IpfcGeneralViewCreateInstructions.Scale
• IpfcGeneralViewCreateInstructions.ViewScale
ThemethodCCpfcGeneralViewCreateInstructions.Create()
createsthe IpfcGeneralViewCreateInstructions dataobjectused for
| creatinggeneraldrawing |     | views. |     |     |     |
| ---------------------- | --- | ------ | --- | --- | --- |
Usethe propertyIpfcGeneralViewCreateInstructions.ViewModel
toassignthesolidmodelto displayin thecreatedgeneraldrawing view.
Usethe property
| IpfcGeneralViewCreateInstructions.SheetNumber |                  |     |                                |     | toassignthe |
| --------------------------------------------- | ---------------- | --- | ------------------------------ | --- | ----------- |
| locationin                                    | a drawingsheetto |     | place thecreatedgeneraldrawing |     | view.       |
Usethe property
| IpfcGeneralViewCreateInstructions.SheetNumber |     |     |     |     | tosetthe |
| --------------------------------------------- | --- | --- | --- | --- | -------- |
numberof thedrawing sheetin whichthegeneraldrawing view iscreated.
ThepropertyIpfcGeneralViewCreateInstructions.Orientation
assignsthe orientationofthe modelinthe generaldrawing viewin theform ofthe
IpfcTransform3Ddataobject.The
transformationmatrixmustonlyconsist of
therotationto be appliedto themodel. It mustnot consistofany displacementor
scalecomponents.Ifnecessary,setthe displacementto {0,0,0} using themethod
| IpfcTransform3D.SetOrigin, |         |     | andremoveanyscalingfactorby |     |     |
| -------------------------- | ------- | --- | --------------------------- | --- | --- |
| normalizingthe             | matrix. |     |                             |     |     |
Usethe propertyIpfcGeneralViewCreateInstructions.Exploded
| tosetthe | createdgeneraldrawing |     | viewto | be an explodedview. |     |
| -------- | --------------------- | --- | ------ | ------------------- | --- |
Usethe propertyIpfcGeneralViewCreateInstructions.Scale to
assigna scaleto thecreatedgeneral drawingview.This valueis optional,if not
| assigned,the | defaultdrawingscale |     | isused. |     |     |
| ------------ | ------------------- | --- | ------- | --- | --- |
Usethe propertyIpfcGeneralViewCreateInstructions.ViewScale
toassignascale tothe createdgeneraldrawing view.This valueis optional,if not
| assigned,the | defaultdrawingscale |     | isused. |                                 |     |
| ------------ | ------------------- | --- | ------- | ------------------------------- | --- |
| 114          |                     |     |         | Creo®ParametricVBAPIUser’sGuide |     |

| Projected | Drawing | Views |     |     |     |
| --------- | ------- | ----- | --- | --- | --- |
TheinterfaceIpfcProjectionViewCreateInstructions
contains
| detailsonhowto | creategeneraldrawing |     | views. |     |     |
| -------------- | -------------------- | --- | ------ | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcProjectionViewCreateInstructions.Create()
• IpfcProjectionViewCreateInstructions.ParentView
• IpfcProjectionViewCreateInstructions.Location
• IpfcProjectionViewCreateInstructions.Exploded
ThemethodCCpfcProjectionViewCreateInstructions.Create()
IpfcProjectionViewCreateInstructions
| createsthe                  |     |        |     |     | dataobjectused |
| --------------------------- | --- | ------ | --- | --- | -------------- |
| forcreatingprojecteddrawing |     | views. |     |     |                |
Usethe property
| IpfcProjectionViewCreateInstructions.ParentView |                         |     |       |     | to assign |
| ----------------------------------------------- | ----------------------- | --- | ----- | --- | --------- |
| theparentview                                   | for theprojecteddrawing |     | view. |     |           |
Usethe property
| IpfcProjectionViewCreateInstructions.Location |     |     |     |     | toassignthe |
| --------------------------------------------- | --- | --- | --- | --- | ----------- |
locationof theprojecteddrawing view.This locationdetermineshowthedrawing
| viewwillbe | oriented. |     |     |     |     |
| ---------- | --------- | --- | --- | --- | --- |
Usethe property
| IpfcProjectionViewCreateInstructions.Exploded |                              |        |                     |                 | tosetthe |
| --------------------------------------------- | ---------------------------- | ------ | ------------------- | --------------- | -------- |
| createdprojecteddrawing                       |                              | viewto | be an explodedview. |                 |          |
| Example:Creating                              | DrawingViews                 |        |                     |                 |          |
| Thesamplecodein                               | thefilepfcDrawingExamples.vb |        |                     | locatedat<creo_ |          |
weblink_loadpoint>/weblinkexamples/jscript<creo_vbapi_
| loadpoint>/vbapi_examplesaddsanewsheetto |                 |       |     | adrawing | andcreates |
| ---------------------------------------- | --------------- | ----- | --- | -------- | ---------- |
| threeviewsof                             | aselectedmodel. |       |     |          |            |
| Obtaining                                | Drawing         | Views |     |          |            |
MethodsandPropertyIntroduced:
• IpfcSelection.SelView2D
• IpfcpfcModel2D.List2DViews()
• IpfcModel2D.GetViewByName()
• IpfcModel2D.GetViewDisplaying()
• IpfcSheetOwner.GetSheetBackgroundView()
Drawings 115

ThepropertyIpfcSelection.SelView2D returns theselecteddrawing view
(ifthe userselectedan item froma drawingview). Itreturns anullvalueif the
| selectiondoesnotcontain               | adrawing | view. |               |     |
| ------------------------------------- | -------- | ----- | ------------- | --- |
| ThemethodIpfcpfcModel2D.List2DViews() |          | lists | andreturnsthe |     |
drawingviews found.This methoddoesnot includethedrawing sheetbackground
| viewsreturnedby themethod |     |     |     |     |
| ------------------------- | --- | --- | --- | --- |
IpfcSheetOwner.GetSheetBackgroundView().
ThemethodIpfcModel2D.GetViewByName()
|     |     | returnsthedrawing |     | view |
| --- | --- | ----------------- | --- | ---- |
basedon thename.This methodreturns anullvalueif thespecifiedviewdoes not
exist.
| ThemethodIpfcModel2D.GetViewDisplaying() |     |     | returnsthe | drawing |
| ---------------------------------------- | --- | --- | ---------- | ------- |
viewthatdisplaysadimension.This methodreturns anullvalueif thedimension
| isnotdisplayedin thedrawing. |     |     |     |     |
| ---------------------------- | --- | --- | --- | --- |
Note
| This methodworksfor                              | solidand | drawingdimensions. |     |             |
| ------------------------------------------------ | -------- | ------------------ | --- | ----------- |
| ThemethodIpfcSheetOwner.GetSheetBackgroundView() |          |                    |     | returns the |
drawingsheetbackgroundviews.
| Drawing View | Information |     |     |     |
| ------------ | ----------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcChild.DBParent
• IpfcView2D.GetSheetNumber()
• IpfcView2D.IsBackground
• IpfcView2D.GetModel()
• IpfcView2D.Scale()
• IpfcView2D.GetIsScaleUserdefined()
• IpfcView2D.Outline
• IpfcView2D.GetLayerDisplayStatus()
• IpfcView2D.IsViewdisplayLayerDependent
• IpfcView2D.Display
• IpfcView2D.GetTransform()
• IpfcView2D.Name
• IpfcView2D.GetSimpRep()
| 116 |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | ------------------------------- | --- | --- |

| TheinheritedpropertyIpfcChild.DBParent, |     |     |     |     |     |     | whencalledon | a   |     |
| --------------------------------------- | --- | --- | --- | --- | --- | --- | ------------ | --- | --- |
IpfcView2Dobject,providesthe drawingmodelwhichownsthe specified
| drawingview.The |     | returnvalue |     | ofthemethod |     | canbe | downcastto | a   |     |
| --------------- | --- | ----------- | --- | ----------- | --- | ----- | ---------- | --- | --- |
IpfcModel2Dobject.
Note
| The                                  | propertyIpfcChild.OIdis |     |              |     | reservedforinternaluse. |     |                         |     |     |
| ------------------------------------ | ----------------------- | --- | ------------ | --- | ----------------------- | --- | ----------------------- | --- | --- |
| ThemethodIpfcView2D.GetSheetNumber() |                         |     |              |     |                         |     | returnsthesheetnumberof |     |     |
| thesheetthatcontainsthe              |                         |     | drawingview. |     |                         |     |                         |     |     |
ThepropertyIpfcView2D.IsBackground
|            |        |                 |     |     |     | returns    | avaluethatindicates |     |     |
| ---------- | ------ | --------------- | --- | --- | --- | ---------- | ------------------- | --- | --- |
| whetherthe | viewis | abackgroundview |     |     | ora | modelview. |                     |     |     |
ThemethodIpfcView2D.GetModel() returns thesolidmodeldisplayedin
| thedrawing | view. |     |     |     |     |     |     |     |     |
| ---------- | ----- | --- | --- | --- | --- | --- | --- | --- | --- |
ThepropertyIpfcView2D.Scale() returnsthescaleof thedrawing view.
| ThemethodIpfcView2D.GetIsScaleUserdefined() |     |     |        |     |     |     |     | specifiesif | the |
| ------------------------------------------- | --- | --- | ------ | --- | --- | --- | --- | ----------- | --- |
| drawinghasauser-defined                     |     |     | scale. |     |     |     |     |             |     |
ThepropertyIpfcView2D.Outline returnstheposition oftheview in the
| sheetinworld | units. |     |     |     |     |     |     |     |     |
| ------------ | ------ | --- | --- | --- | --- | --- | --- | --- | --- |
ThemethodIpfcView2D.GetLayerDisplayStatus() returnsthedisplay
| statusofthespecifiedlayerin |     |     |     | thedrawing |     | view. |     |     |     |
| --------------------------- | --- | --- | --- | ---------- | --- | ----- | --- | --- | --- |
ThepropertyIpfcView2D.Display returnsan outputstructure thatdescribes
thedisplaysettingsofthe drawingview.The fieldsin thestructure areasfollows:
• Style—Whetherto displayaswireframe,hidden lines,no hidden lines,or
shaded
| • TangentStyle—Linestyleusedfortangent |     |     |     |               |     | edges |     |     |     |
| -------------------------------------- | --- | --- | --- | ------------- | --- | ----- | --- | --- | --- |
| • CableStyle—Linestyleusedto           |     |     |     | displaycables |     |       |     |     |     |
• RemoveQuiltHiddenLines—Whetherornotto applyhidden-line-removal to
quilts
| • ShowConceptModel—Whetherornotto |     |     |     |     |     | displaytheskeleton |     |     |     |
| --------------------------------- | --- | --- | --- | --- | --- | ------------------ | --- | --- | --- |
• ShowWeldXSection—Whetherornotto includewelds in thecross-section
ThemethodIpfcView2D.GetTransform() returnsamatrixthatdescribes
thetransformbetween 3Dsolidcoordinatesand2Dworld unitsforthatdrawing
view. Thetransformationmatrix isacombinationof thefollowingfactors:
| • Thelocationof |           | theview |                    | originwith | respecttothe |            | drawingorigin. |     |     |
| --------------- | --------- | ------- | ------------------ | ---------- | ------------ | ---------- | -------------- | --- | --- |
| • Thescale      | oftheview |         | unitswithrespectto |            |              | thedrawing | units          |     |     |
| Drawings        |           |         |                    |            |              |            |                |     | 117 |

• Therotation ofthe modelwithrespectto thedrawing coordinatesystem.
ThepropertyIpfcView2D.Name
|     |     |     | returnsthe | nameofthe | specifiedviewin | the |
| --- | --- | --- | ---------- | --------- | --------------- | --- |
drawing.
Thesimplifiedrepresentationsof assemblyandpart canbe usedas drawing
| modelsto | creategeneralviews. | Usethe | method |     |     |     |
| -------- | ------------------- | ------ | ------ | --- | --- | --- |
IpfcView2D.GetSimpRep()toretrieve thesimplifiedrepresentationfor the
| specifiedview                            | in thedrawing.               |           |      |         |                   |     |
| ---------------------------------------- | ---------------------------- | --------- | ---- | ------- | ----------------- | --- |
| Example:                                 | Listing                      | the Views | in a | Drawing |                   |     |
| Thesamplecodein                          | thefilepfcDrawingExamples.vb |           |      |         | locatedat<creo_   |     |
| vbapi_loadpoint>/vbapi_examplescreatesan |                              |           |      |         | informationwindow |     |
aboutallthe viewsin adrawing. Theinformation isplacedin anexternalbrowser
window.
| Drawing | Views | Operations |     |     |     |     |
| ------- | ----- | ---------- | --- | --- | --- | --- |
MethodsIntroduced:
• IpfcView2D.Translate()
• IpfcView2D.Delete()
• IpfcView2D.Regenerate()
• IpfcView2D.SetLayerDisplayStatus()
| ThemethodIpfcView2D.Translate() |     |         |     | movesthe | drawingview | by the |
| ------------------------------- | --- | ------- | --- | -------- | ----------- | ------ |
| specifiedtransformation         |     | vector. |     |          |             |        |
ThemethodIpfcView2D.Delete() deletes aspecifieddrawingview. Setthe
DeleteChildren parameterto trueto deletethechildren oftheview. Setthis
parameterto falseornullto preventdeletionoftheview ifithaschildren.
ThemethodIpfcView2D.Regenerate() erasesthe displayedviewof the
currentobject,regeneratestheview fromthe currentdrawing, andredisplaysthe
view.
| ThemethodIpfcView2D.SetLayerDisplayStatus() |            |              |     |     | setsthedisplay |     |
| ------------------------------------------- | ---------- | ------------ | --- | --- | -------------- | --- |
| statusforthe                                | layerinthe | drawingview. |     |     |                |     |
| Drawing                                     | Dimensions |              |     |     |                |     |
Thissectiondescribesthe VBAPI methodsthatgiveaccessto thetypesof
dimensionsthatcanbe createdin thedrawing mode.Theydo notapplyto
dimensionscreatedin thesolidmode,eitherthosecreatedautomaticallyas aresult
offeaturecreation,or referencedimension createdin asolid. Adrawing
dimensionora referencedimension shownin adrawing isrepresented bythe
interfaceIpfcDimension2D.
| 118 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

Obtaining Drawing Dimensions
MethodsandPropertyIntroduced:
• IpfcModelItemOwner.ListItems()
• IpfcModelItemOwner.GetItemById()
• IpfcSelection.SelItem
ThemethodIpfcModelItemOwner.ListItems() returns alist ofdrawing
dimensionsspecifiedbythe parameterType orreturnsnullif nodrawing
dimensionsofthe specifiedtypearefound. Thismethod listsonly those
dimensionscreatedin thedrawing.
Thevalues oftheparameter Typefor thedrawing dimensionsare:
• ITEM_DIMENSION—Dimension
• ITEM_REF_DIMENSION—Reference dimension
SettheparameterTypeto thetypeofdrawing dimensionto retrieve.Ifthis
parameterissetto null, thenall thedimensions inthe drawing arelisted.
ThemethodIpfcModelItemOwner.GetItemById() returns adrawing
dimensionbasedonthe typeand theintegeridentifier.The methodreturns only
thosedimensions createdin thedrawing. Itreturnsa nullifadrawing dimension
withthespecifiedattributes isnotfound.
ThepropertyIpfcSelection.SelItem returns thevalueof theselected
drawingdimension.
Creating Drawing Dimensions
MethodsIntroduced:
• CCpfcDrawingDimCreateInstructions.Create()
• IpfcModel2D.CreateDrawingDimension()
• CCpfcEmptyDimensionSense.Create()
• CCCpfcPointDimensionSense.Create()
• CCpfcSplinePointDimensionSense.Create()
• CCpfcTangentIndexDimensionSense.Create()
• CCpfcLinAOCTangentDimensionSense.Create()
• CCpfcAngleDimensionSense.Create()
• CCpfcPointToAngleDimensionSense.Create()
ThemethodCCpfcDrawingDimCreateInstructions.Create()
createsan instructionsobjectthatdescribeshowto createadrawing dimension
usingthe methodIpfcModel2D.CreateDrawingDimension().
Drawings 119

Theparametersoftheinstructionobjectare:
• Attachments—Theentitiesthatthe dimensionisattached to.The selections
| shouldincludethe |     | drawingmodelview. |     |     |     |     |
| ---------------- | --- | ----------------- | --- | --- | --- | --- |
• IsRefDimension—Trueifthe dimensionisa referencedimension,otherwise
nullorfalse.
• OrientationHint—Describestheorientationof thedimensions incases where
| thiscannotbe | deducedfrom |     | theattachmentsthemselves. |     |     |     |
| ------------ | ----------- | --- | ------------------------- | --- | --- | --- |
• Senses—Gives moreinformationabouthowthedimension attachesto the
entity,i.e.,to what partof theentityandin whatdirection thedimensionruns.
| Thetypesof | dimensionsensesare |     | asfollows: |     |     |     |
| ---------- | ------------------ | --- | ---------- | --- | --- | --- |
○ EpfcDIMSENSE_NONE
○ EpfcDIMSENSE_POINT
EpfcDIMSENSE_SPLINE_PT
○
○ EpfcDIMSENSE_TANGENT_INDEX
○ EpfcDIMSENSE_LINEAR_TO_ARC_OR_CIRCLE_TANGENT
EpfcDIMSENSE_ANGLE
○
○ EpfcDIMSENSE_POINT_TO_ANGLE
| • TextLocation—Thelocationofthedimension      |     |     |     |     | text,in | world units. |
| --------------------------------------------- | --- | --- | --- | --- | ------- | ------------ |
| ThemethodIpfcModel2D.CreateDrawingDimension() |     |     |     |     |         | createsa     |
dimensionin thedrawing basedontheinstructionsdataobjectthatcontains
informationneededtoplace thedimension.It takesasinputan arrayof
pfcSelectionobjectsandan arrayofpfcDimensionSense structuresthat
describetherequiredattachments. Themethod returnsthecreated drawing
dimension.
| ThemethodCCpfcEmptyDimensionSense.Create() |     |     |     |     |     | createsanew |
| ------------------------------------------ | --- | --- | --- | --- | --- | ----------- |
dimensionsenseassociated withthe typeDIMSENSENONE. The sensefieldis
setto TypeIn thiscase noinformation suchaslocationor directionisneededto
describetheattachmentpoints.For example,if thereisa singleattachmentwhich
isastraightline, thedimension isthelengthof thestraightline. Iftheattachments
| aretwoparallellines,the                     |     | dimensionisthe |     | distancebetween |     | them.       |
| ------------------------------------------- | --- | -------------- | --- | --------------- | --- | ----------- |
| ThemethodCCCpfcPointDimensionSense.Create() |     |                |     |                 |     | createsanew |
dimensionsenseassociated withthe typeDIMSENSEPOINT whichspecifiesthe
partofthe entitytowhich thedimensionis attached.The sensefield issetto the
| valueofthe parameterPointType. |     |               |                |          |                                 |     |
| ------------------------------ | --- | ------------- | -------------- | -------- | ------------------------------- | --- |
| Thepossiblevaluesof            |     | PointTypeare: |                |          |                                 |     |
| • EpfcDIMPOINT_END1—           |     |               | Thefirst       | endofthe | entity                          |     |
| • EpfcDIMPOINT_END2—The        |     |               | secondendofthe |          | entity                          |     |
| 120                            |     |               |                |          | Creo®ParametricVBAPIUser’sGuide |     |

| • EpfcDIMPOINT_CENTER—The |     |     |     |     | centerof | anarc | orcircle |     |     |     |
| ------------------------- | --- | --- | --- | --- | -------- | ----- | -------- | --- | --- | --- |
• EpfcDIMPOINT_NONE—Noinformation suchaslocationor directionofthe
attachmentisspecified.This issimilar to settingthePointTypeto DIMSENSE
NONE.
| • EpfcDIMPOINT_MIDPOINT—The                      |     |     |     |     | midpoint |     | oftheentity |     |          |     |
| ------------------------------------------------ | --- | --- | --- | --- | -------- | --- | ----------- | --- | -------- | --- |
| ThemethodCCpfcSplinePointDimensionSense.Create() |     |     |     |     |          |     |             |     | createsa |     |
dimensionsenseassociated withthe typeDIMSENSE_SPLINE_PT. Thismeans
| thatthe attachmentisto                            |     |     | apointon      |     | aspline.The  | sensefieldis |     |     | setto |          |
| ------------------------------------------------- | --- | --- | ------------- | --- | ------------ | ------------ | --- | --- | ----- | -------- |
| SplinePointIndexi.e.,                             |     |     | theindexofthe |     | splinepoint. |              |     |     |       |          |
| ThemethodCCpfcTangentIndexDimensionSense.Create() |     |     |               |     |              |              |     |     |       | createsa |
newdimensionsenseassociated withthe typeDIMSENSE_TANGENT_INDEX.
Theattachmentisto atangent oftheentity, whichisan arcor acircle. Thesense
| fieldissetto                                       | TangentIndex, |     |     | i.e.,the | indexofthetangent |     |     | ofthe entity. |     |         |
| -------------------------------------------------- | ------------- | --- | --- | -------- | ----------------- | --- | --- | ------------- | --- | ------- |
| ThemethodCCpfcLinAOCTangentDimensionSense.Create() |               |     |     |          |                   |     |     |               |     | creates |
anewdimensionsense associatedwith thetypeDIMSENSE_LINEAR_TO_ARC_
OR_CIRCLE_TANGENT.The dimensionisthe perpendiculardistancebetween
thealineanda tangentto anarc oracirclethatis parallelto the line.The sense
| fieldissetto        | thevalueof |     | theparameterTangentType. |     |      |     |     |     |     |     |
| ------------------- | ---------- | --- | ------------------------ | --- | ---- | --- | --- | --- | --- | --- |
| Thepossiblevaluesof |            |     | TangentType              |     | are: |     |     |     |     |     |
• EpfcDIMLINAOCTANGENT_LEFT0—Thetangent isto theleftofthe line,
| andison                           | thesameside,of |     |     | thecenterofthe |     | arcorcircle, |           | astheline. |            |     |
| --------------------------------- | -------------- | --- | --- | -------------- | --- | ------------ | --------- | ---------- | ---------- | --- |
| • EpfcDIMLINAOCTANGENT_RIGHT0—The |                |     |     |                |     |              | tangentis | tothe      | rightofthe |     |
line,and ison thesameside,ofthe centerof thearcor circle,as theline.
EpfcDIMLINAOCTANGENT_LEFT1—Thetangent
| •       |             |     |           |     |       |     |     | isto theleftofthe |     | line, |
| ------- | ----------- | --- | --------- | --- | ----- | --- | --- | ----------------- | --- | ----- |
| andison | theopposite |     | sideofthe |     | line. |     |     |                   |     |       |
• EpfcDIMLINAOCTANGENT_RIGHT1— The tangentisto the rightofthe
| line,and                                   | ison | theoppositeside |     |     | oftheline. |     |     |             |     |     |
| ------------------------------------------ | ---- | --------------- | --- | --- | ---------- | --- | --- | ----------- | --- | --- |
| ThemethodCCpfcAngleDimensionSense.Create() |      |                 |     |     |            |     |     | createsanew |     |     |
dimensionsenseassociated withthe typeDIMSENSE_ANGLE. The dimensionis
theanglebetweentwo straightentities.The sensefieldis settothe valueof the
parameterAngleOptions.
| Thepossiblevaluesof |     |       | AngleOptionsare: |     |                   |     |               |     |           |     |
| ------------------- | --- | ----- | ---------------- | --- | ----------------- | --- | ------------- | --- | --------- | --- |
| IsFirst—Is          |     |       | TRUEif           |     |                   |     |               |     |           |     |
| •                   |     | setto |                  |     | theangledimension |     | startsfromthe |     | specified |     |
entityina counterclockwisedirection.Issetto FALSEif thedimensionends
atthespecifiedentity.The valueis TRUEforoneentityandFALSE for the
otherentityformingtheangle.
• ShouldFlip—Ifthe valueofShouldFlip isFALSE,andthe directionof
thespecifiedentityisaway from thevertexof theangle,thenthedimension
attachesdirectlyto theentity. Ifthedirection oftheentityis awayfrom the
| Drawings |     |     |     |     |     |     |     |     |     | 121 |
| -------- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |

vertexof theangle,thenthedimension isattachedto thea witnessline. The
witnesslineisin linewith theentitybutin thedirectionopposite to thevertex
oftheangle. Ifthevalueof ShouldFlipisTRUEthentheabove casesare
reversed.
ThemethodCCpfcPointToAngleDimensionSense.Create() createsa
newdimensionsenseassociated withthe typeDIMSENSE_POINT_TO_ANGLE.
Thedimensionis theanglebetween alineentityandthetangent to acurvedentity.
Thecurveattachmentisof thetypeDIMSENSE_POINT_TO_ANGLE andtheline
attachmentis ofthetypeDIMSENSEPOINT.In thiscase boththe angleandthe
angle_sensefieldsmustbe set. Thefieldsense showswhichendof the
curvethedimensionis attachedtoand thefieldangle_sense showsthe
directioninwhich thedimensionrotatesandto whichsideof thetangent it
attaches.
Drawing Dimensions Information
MethodsandPropertiesIntroduced:
• IpfcDimension2D.IsAssociative
• IpfcDimension2D.GetIsReference()
• IpfcDimension2D.IsDisplayed
• IpfcDimension2D.GetAttachmentPoints()
• IpfcDimension2D.GetDimensionSenses()
• IpfcDimension2D.GetOrientationHint()
• IpfcDimension2D.GetBaselineDimension()
• IpfcDimension2D.Location
• IpfcDimension2D.GetView()
• IpfcDimension2D.GetTolerance()
• IpfcDimension2D.IsToleranceDisplayed
ThepropertyIpfcDimension2D.IsAssociative returnswhetherthe
dimensionorreferencedimensionin adrawing is associative.
ThemethodIpfcDimension2D.GetIsReference() determineswhether
thedrawing dimensionisa referencedimension.
ThepropertyIpfcDimension2D.IsDisplayed determineswhetherthe
dimensionwillbedisplayedin thedrawing.
ThemethodIpfcDimension2D.GetAttachmentPoints() returnsa
sequenceofattachmentpoints.The dimensionsenses arrayreturned bythe
methodIpfcDimension2D.GetDimensionSenses() givesmore
informationonhowtheseattachmentsareinterpreted.
122 Creo®ParametricVBAPIUser’sGuide

| ThemethodIpfcDimension2D.GetDimensionSenses() |     |     |     | returnsa |
| --------------------------------------------- | --- | --- | --- | -------- |
sequenceofdimension senses,describinghowthedimension isattachedto each
| attachmentreturned | bythemethod |     |     |     |
| ------------------ | ----------- | --- | --- | --- |
IpfcDimension2D.GetAttachmentPoints().
| ThemethodIpfcDimension2D.GetOrientationHint() |     |     |     | returnsthe |
| --------------------------------------------- | --- | --- | --- | ---------- |
orientationhintforplacing thedrawing dimensions.Theorientation hint
determineshowCreo Parametricwillorientthe dimensionwith respectto the
attachmentpoints.
Note
This methodsdescribed aboveareapplicableonly fordimensions createdin
thedrawing mode.It doesnotsupportdimensions createdatintersection
| pointsof                                        | entities. |     |     |           |
| ----------------------------------------------- | --------- | --- | --- | --------- |
| ThemethodIpfcDimension2D.GetBaselineDimension() |           |     |     | returnsan |
ordinatebaseline drawingdimension. Itreturnsa nullvalue ifthe dimensionisnot
| anordinate | dimension. |     |     |     |
| ---------- | ---------- | --- | --- | --- |
Note
The methodupdatesthedisplayofthe dimensiononly ifitiscurrently
displayed.
ThepropertyIpfcDimension2D.Location returnsthe placementlocation
ofthedimension.
| ThemethodIpfcDimension2D.GetView() |     |     | returnsthedrawing | view in |
| ---------------------------------- | --- | --- | ----------------- | ------- |
whichthedimensionis displayed.This methodappliesto dimensionsstoredin the
| solidor inthe | drawing. |     |     |     |
| ------------- | -------- | --- | --- | --- |
ThemethodIpfcDimTolerance
|     |     | object.Anullvalue | indicatesanominal |     |
| --- | --- | ----------------- | ----------------- | --- |
tolerance.
Usethe propertyIpfcDimension2D.IsToleranceDisplayed
determineswhetheror notthedimension’s toleranceis displayedin the drawing.
| Drawing | Dimensions | Operations |     |     |
| ------- | ---------- | ---------- | --- | --- |
MethodsIntroduced:
• IpfcDimension2D.ConvertToLinear()
• IpfcDimension2D.ConvertToOrdinate()
• IpfcDimension2D.ConvertToBaseline()
Drawings 123

• IpfcDimension2D.SwitchView()
• IpfcDimension2D.SetTolerance()
• IpfcDimension2D.EraseFromModel2D()
• IpfcModel2D.SetViewDisplaying()
ThemethodIpfcDimension2D.ConvertToLinear()
convertsan
ordinatedrawing dimensionto alineardrawing dimension.Thedrawing
| containingthe | dimensionmustbe |     |     | displayed. |     |     |     |     |
| ------------- | --------------- | --- | --- | ---------- | --- | --- | --- | --- |
ThemethodIpfcDimension2D.ConvertToOrdinate() convertsa linear
| drawingdimension                             |     | toan | ordinate | baselinedimension. |     |     |     |           |
| -------------------------------------------- | --- | ---- | -------- | ------------------ | --- | --- | --- | --------- |
| ThemethodIpfcDimension2D.ConvertToBaseline() |     |      |          |                    |     |     |     | convertsa |
locationon alineardrawing dimensionto an ordinatebaseline dimension.The
| methodreturnsthenewlycreated |     |     |     | baselinedimension. |     |     |     |     |
| ---------------------------- | --- | --- | --- | ------------------ | --- | --- | --- | --- |
Note
The methodupdatesthedisplayofthe dimensiononly ifitiscurrently
displayed.
| ThemethodIpfcDimension2D.SwitchView()   |     |             |                   |              |     | changesthe |            | viewwherea |
| --------------------------------------- | --- | ----------- | ----------------- | ------------ | --- | ---------- | ---------- | ---------- |
| dimensioncreatedin                      |     | thedrawing  |                   | isdisplayed. |     |            |            |            |
| ThemethodIpfcDimension2D.SetTolerance() |     |             |                   |              |     |            | assignsthe | upperand   |
| lowertolerance                          |     | limitsofthe | drawingdimension. |              |     |            |            |            |
ThemethodIpfcDimension2D.EraseFromModel2D()
permanently
| erasesthe                                    | dimensionfrom |                              | thedrawing.             |     |               |     |                 |              |
| -------------------------------------------- | ------------- | ---------------------------- | ----------------------- | --- | ------------- | --- | --------------- | ------------ |
| ThemethodIpfcModel2D.SetViewDisplaying()     |               |                              |                         |     |               |     | changestheview  |              |
| wherea dimensioncreatedin                    |               |                              | asolidmodelisdisplayed. |     |               |     |                 |              |
| Example:                                     | Command       |                              | Creation                |     | of Dimensions |     | from            | Model        |
| Datum                                        | Points        |                              |                         |     |               |     |                 |              |
| Thesamplecodein                              |               | thefilepfcDrawingExamples.vb |                         |     |               |     | locatedat<creo_ |              |
| vbapi_loadpoint>/vbapi_examplesshowsacommand |               |                              |                         |     |               |     |                 | whichcreates |
verticaland horizontalordinate dimensionsfrom eachdatumpointin amodelin a
| drawingview | to  | aselectedcoordinatesystem |     |     | datum. |     |     |     |
| ----------- | --- | ------------------------- | --- | --- | ------ | --- | --- | --- |
| Drawing     |     | Tables                    |     |     |        |     |     |     |
Adrawingtable in theVBAPI isrepresented bythe IpfcTable.It isachild of
| theIpfcModelItem |     | interface. |     |     |     |                                 |     |     |
| ---------------- | --- | ---------- | --- | --- | --- | ------------------------------- | --- | --- |
| 124              |     |            |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |

Somedrawingtable methodsoperateon specificrows orcolumns.The row and
columnnumbersinthe VBAPI beginwith 1andrangeup tothe totalnumberof
rowsor columnsin thetable.Somedrawing table methodsoperateon specific
tablecells.TheinterfaceIpfcTableCell isusedto represent adrawing table
cell.
| Creating | Drawing | Cells |     |     |
| -------- | ------- | ----- | --- | --- |
MethodIntroduced:
• CCpfcTableCell.Create()
| ThemethodCCpfcTableCell.Create() |                  |            |                 | createstheIpfcTableCell |
| -------------------------------- | ---------------- | ---------- | --------------- | ----------------------- |
| objectrepresentinga              | cellin           | thedrawing | table.          |                         |
| Somedrawingtable                 | methodsoperateon |            | specificdrawing | segment.A               |
multisegmenteddrawing tablecontains 2ormoreareasin thedrawing. Inserting
ordeletingrows inone segmentofthe tablecan affectthecontents ofother
segments.Tablesegments arenumberedbeginningwith 0.Ifthetable hasonly a
| singlesegment,use0asthe |         | segmentid | in therelevantmethods. |       |
| ----------------------- | ------- | --------- | ---------------------- | ----- |
| Selecting               | Drawing | Tables    | and                    | Cells |
MethodsandPropertiesIntroduced:
• IpfcBaseSession.Select()
• IpfcSelection.SelItem
• IpfcSelection.SelTableCell
• IpfcSelection.SelTableSegment
Tablesmaybe selectedusing themethod IpfcBaseSession.Select().
Passthe filterdwg_table toselectan entiretableandthe filtertable_cell
| toprompttheuser | toselectaparticular |     | tablecell. |     |
| --------------- | ------------------- | --- | ---------- | --- |
ThepropertyIpfcSelection.SelItem returns theselectedtablehandle. It
| isamodelitem | thatcan | be castto aIpfcTable |     | object. |
| ------------ | ------- | -------------------- | --- | ------- |
ThepropertyIpfcSelection.SelTableCell returnsthe rowand column
| indicesoftheselectedtable |     | cell. |     |     |
| ------------------------- | --- | ----- | --- | --- |
ThepropertyIpfcSelection.SelTableSegment returnsthetable segment
identifierfor theselectedtablecell.Ifthe tableconsists ofasingle segment,this
| methodreturnstheidentifier |         | 0.     |     |     |
| -------------------------- | ------- | ------ | --- | --- |
| Creating                   | Drawing | Tables |     |     |
MethodsIntroduced:
• CCpfcTableCreateInstructions.Create()
Drawings 125

• IpfcTableOwner.CreateTable()
ThemethodCCpfcTableCreateInstructions.Create()
createsthe
IpfcTableCreateInstructionsdataobjectthatdescribeshowto
constructa newtableusing themethodIpfcTableOwner.CreateTable().
Theparametersoftheinstructionsdataobjectare:
• Origin—Thisparameterstores athreedimensional pointspecifyingthe
locationofthetable origin.The originis thepositionof thetopleftcornerof
thetable.
| • RowHeights—Specifiesthe        |     | heightof | eachrowof       | thetable.   |     |
| -------------------------------- | --- | -------- | --------------- | ----------- | --- |
| • ColumnData—Specifiesthewidthof |     |          | eachcolumnofthe | tableandits |     |
justification.
• SizeTypes—Indicates thescaleused tomeasure thecolumnwidthandrow
heightofthetable.
| ThemethodIpfcTableOwner.CreateTable() |     |     |     | createsa tablein | the |
| ------------------------------------- | --- | --- | --- | ---------------- | --- |
drawingspecifiedby theIpfcTableCreateInstructionsdataobject.
| Retrieving | Drawing | Tables |     |     |     |
| ---------- | ------- | ------ | --- | --- | --- |
MethodsIntroduced
• CCpfcTableRetrieveInstructions.Create()
• IpfcTableRetrieveInstructions.FileName
• IpfcTableRetrieveInstructions.Path
• IpfcTableRetrieveInstructions.Version
• IpfcTableRetrieveInstructions.Position
• IpfcTableRetrieveInstructions.ReferenceSolid
• IpfcTableRetrieveInstructions.ReferenceRep
• IpfcTableOwner.RetrieveTable()
• IpfcTableOwner.RetrieveTableByOrigin()
| ThemethodIpfcTableOwner.RetrieveTable() |     |     |     | retrieves atable |     |
| --------------------------------------- | --- | --- | --- | ---------------- | --- |
specifiedby thepfcTableRetrieveInstructions dataobjectfrom afile
onthedisk.It returnsthe retrievedtable.Thedataobjectcontains informationon
| thetableto | retrieve andisreturned | bythemethod |     |     |     |
| ---------- | ---------------------- | ----------- | --- | --- | --- |
CCpfcTableRetrieveInstructions.Create().
ThemethodCCpfcTableRetrieveInstructions.Create()
createsthe
IpfcTableRetrieveInstructionsdataobjectthatdescribeshowto
| retrieveadrawing | table usingthe | methods |     |                                 |     |
| ---------------- | -------------- | ------- | --- | ------------------------------- | --- |
| 126              |                |         |     | Creo®ParametricVBAPIUser’sGuide |     |

| IpfcTableOwner.RetrieveTable()          |     |     | and |                  |     |
| --------------------------------------- | --- | --- | --- | ---------------- | --- |
| IpfcTableOwner.RetrieveTableByOrigin(). |     |     |     | Themethodreturns | the |
createdinstructionsdataobject.
Theparametersoftheinstructionobjectare:
| • FileName—Nameof |     | thefilecontainingthedrawing |     | table. |     |
| ----------------- | --- | --------------------------- | --- | ------ | --- |
• Position—Coordinatesof thepointon thedrawing sheet, wheretheretrieved
tablemustbe placed.You mustspecify thevalue in screencoordinates.
You canalsosetthe parametersfor IpfcTableRetrieveInstructions
| dataobjectusing | thefollowingproperty: |     |     |     |     |
| --------------- | --------------------- | --- | --- | --- | --- |
• IpfcTableRetrieveInstructions.FileName—Sets thenameof
| thedrawing | table.Youmustnotspecify |     | theextension. |     |     |
| ---------- | ----------------------- | --- | ------------- | --- | --- |
• IpfcTableRetrieveInstructions.Path—Setsthe pathtothe
| drawingtable | file.The | pathmustbespecifiedrelativeto |     | theworking |     |
| ------------ | -------- | ----------------------------- | --- | ---------- | --- |
directory.
| • IpfcTableRetrieveInstructions.Version—Sets |     |     |     | theversion | of  |
| -------------------------------------------- | --- | --- | --- | ---------- | --- |
thedrawing tablethatmustbe retrieved.IfyouspecifyNULL, thelatest
| versionof | thedrawing | tableis retrieved. |     |     |     |
| --------- | ---------- | ------------------ | --- | --- | --- |
• IpfcTableRetrieveInstructions.Position—Sets the
coordinatesof thepointon thedrawing sheet, wherethetable mustbe placed.
| You mustspecifythe                                  |     | valuein screencoordinates. |     |     |     |
| --------------------------------------------------- | --- | -------------------------- | --- | --- | --- |
| • IpfcTableRetrieveInstructions.ReferenceSolid—Sets |     |                            |     |     | the |
modelfromwhichdatamustbe copiedintothe drawingtable. Ifthisargument
| ispassedasNULL,                                   |     | an emptytableiscreated. |     |     |     |
| ------------------------------------------------- | --- | ----------------------- | --- | --- | --- |
| • IpfcTableRetrieveInstructions.ReferenceRep—Sets |     |                         |     |     | the |
handletothe simplifiedrepresentation ina solid,from whichdatamustbe
copiedintothe drawingtable. Ifthisargumentispassedas NULL,and the
argument solidis notNULL, thendatafrom thesolidmodeliscopiedintothe
drawingtable
| ThemethodIpfcTableOwner.RetrieveTable() |     |     |     | retrieves atable |     |
| --------------------------------------- | --- | --- | --- | ---------------- | --- |
specifiedby theIpfcTableRetrieveInstructions dataobjectfroma file
onthedisk.It returnsthe retrievedtable.Theupper-leftcorner ofthetable is
| placedon | thedrawing | sheetattheposition | specifiedbythe |     |     |
| -------- | ---------- | ------------------ | -------------- | --- | --- |
pfcTableRetrieveInstructionsdataobject.
| ThemethodIpfcTableOwner.RetrieveTableByOrigin() |     |     |     | also |     |
| ----------------------------------------------- | --- | --- | --- | ---- | --- |
retrievesatable specifiedbythe IpfcTableRetrieveInstructionsdata
objectfrom afileon thedisk.The originofthe tableis placedonthe drawing
sheetatthe positionspecifiedbythe IpfcTableRetrieveInstructions
dataobject.Tablescanbe createdwith differentorigins byspecifyingthe option
| Direction, | inthe InsertTabledialogbox. |     |     |     |     |
| ---------- | --------------------------- | --- | --- | --- | --- |
| Drawings   |                             |     |     |     | 127 |

| Drawing | Tables | Information |     |     |     |     |
| ------- | ------ | ----------- | --- | --- | --- | --- |
MethodsIntroduced:
• IpfcTableOwner.ListTables()
• IpfcTableOwner.GetTable()
• IpfcTable.GetRowCount()
• IpfcTable.GetColumnCount()
• IpfcTable.CheckIfIsFromFormat()
• IpfcTable.GetRowSize()
• IpfcTable.GetColumnSize()
• IpfcTable.GetText()
• IpfcTable.GetCellNote()
ThemethodIpfcTableOwner.ListTables() returnsasequence oftables
| foundin | themodel. |     |     |     |     |     |
| ------- | --------- | --- | --- | --- | --- | --- |
ThemethodIpfcTableOwner.GetTable()
|     |     |     |     | returnsatable | specifiedbythe |     |
| --- | --- | --- | --- | ------------- | -------------- | --- |
tableidentifier inthe model.It returnsanullvalue ifthetable isnotfound.
ThemethodIpfcTable.GetRowCount() returnsthenumberofrows in the
table.
| ThemethodIpfcTable.GetColumnCount()      |           |     |     | returnsthe | numberof |            |
| ---------------------------------------- | --------- | --- | --- | ---------- | -------- | ---------- |
| columnsin                                | thetable. |     |     |            |          |            |
| ThemethodIpfcTable.CheckIfIsFromFormat() |           |     |     |            | checksif | thedrawing |
tablewas createdusing theformat. Themethodreturns atrue valueifthe table
| wascreated | byapplying | thedrawing | format. |     |     |     |
| ---------- | ---------- | ---------- | ------- | --- | --- | --- |
ThemethodIpfcTable.GetRowSize() returnsthe heightof thedrawing
| tablerow                           | specifiedbythesegment |     | identifier andtherow | number.           |     |     |
| ---------------------------------- | --------------------- | --- | -------------------- | ----------------- | --- | --- |
| ThemethodIpfcTable.GetColumnSize() |                       |     |                      | returnsthewidthof |     | the |
drawingtable columnspecifiedbythe segmentidentifier andthecolumn number.
ThemethodIpfcTable.GetText() returns thesequenceof textin adrawing
DWGTABLE_NORMAL
| tablecell.Setthevalue |     | oftheparameter | Modeto |     |     | toget the |
| --------------------- | --- | -------------- | ------ | --- | --- | --------- |
textas displayedonthe screen.SetittoDWGTABLE_FULL to getsymbolictext,
| whichincludesthenames            |          | ofparameter | referencesinthe | tabletext.               |     |     |
| -------------------------------- | -------- | ----------- | --------------- | ------------------------ | --- | --- |
| ThemethodIpfcTable.GetCellNote() |          |             |                 | returnsthedetailnoteitem |     |     |
| containedin                      | thetable | cell.       |                 |                          |     |     |
| Drawing                          | Tables   | Operations  |                 |                          |     |     |
MethodsIntroduced:
• IpfcTable.Erase()
| 128 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | ------------------------------- | --- | --- |

• IpfcTable.Display()
• IpfcTable.RotateClockwise()
• IpfcTable.InsertRow()
• IpfcTable.InsertColumn()
• IpfcTable.MergeRegion()
• IpfcTable.SubdivideRegion()
• IpfcTable.DeleteRow()
• IpfcTable.DeleteColumn()
• IpfcTable.SetText()
• IpfcTableOwner.DeleteTable()
ThemethodIpfcTable.Erase()
|     |     |     | erasesthespecifiedtable |     |     | temporarilyfrom |     |
| --- | --- | --- | ----------------------- | --- | --- | --------------- | --- |
thedisplay.It stillexistsin thedrawing. Theerasedtable canbe displayedagain
usingthe methodIpfcTable.Display(). Thetable willalso be redisplayed
byawindow repaintoraregeneration ofthe drawing.Use thesemethodsto hidea
tablefrom thedisplaywhileyouaremakingmultiple changesto thetable.
ThemethodIpfcTable.RotateClockwise() rotatesatableclockwiseby
thespecifiedamountofrotation.
| ThemethodIpfcTable.InsertRow() |     |     |     | inserts | anewrow | in thedrawing |     |
| ------------------------------ | --- | --- | --- | ------- | ------- | ------------- | --- |
table.Setthevalue oftheparameterRowHeight tospecify theheightofthe row.
Setthevalueof theparameterInsertAfterRowto specifythe rownumberafter
whichthenewrow hasto be inserted.Specify0to insertanewfirst row.
| ThemethodIpfcTable.InsertColumn() |     |     |     |     | insertsa | newcolumnin | the |
| --------------------------------- | --- | --- | --- | --- | -------- | ----------- | --- |
drawingtable.Set thevalueof theparameter ColumnWidthto specifythe widthof
thecolumn.Setthevalue oftheparameter InsertAfterColumnto specify the
columnnumberafterwhichthenewcolumnhasto be inserted.Specify0 toinsert
anewfirstcolumn.
| ThemethodIpfcTable.MergeRegion() |     |     |     | mergestable |     | cellswithina |     |
| -------------------------------- | --- | --- | --- | ----------- | --- | ------------ | --- |
specifiedrangeof rowsand columnsto forma singlecell.Therangeis a
rectangularregionspecifiedbythe tablecellon theupperleftofthe regionand
| thetablecellon                       | thelowerrightof | theregion. |     |     |                   |     |     |
| ------------------------------------ | --------------- | ---------- | --- | --- | ----------------- | --- | --- |
| ThemethodIpfcTable.SubdivideRegion() |                 |            |     |     | removesmergesfrom |     | a   |
regionoftable cellsthatwerepreviouslymerged. Theregion to removemergesis
specifiedby thetable cellonthe upperleftof theregion andthetable cellonthe
| lowerrightof                    | theregion. |     |     |     |     |     |     |
| ------------------------------- | ---------- | --- | --- | --- | --- | --- | --- |
| ThemethodsIpfcTable.DeleteRow() |            |     |     | and |     |     |     |
IpfcTable.DeleteColumn()deleteany specifiedrow or columnfromthe
| table.Themethods             | alsoremovethetextfrom |     |            | theaffected |          | cells. |     |
| ---------------------------- | --------------------- | --- | ---------- | ----------- | -------- | ------ | --- |
| ThemethodIpfcTable.SetText() |                       |     | setstextin |             | thetable | cell.  |     |
| Drawings                     |                       |     |            |             |          |        | 129 |

Usethe methodIpfcTableOwner.DeleteTable() todelete aspecified
drawingtable fromthe modelpermanently. Thedeleted tablecannotbe displayed
again.
Note
Many oftheabove methodsprovideaparameter RepaintIf thisissetto true
| thetable | willberepainted |     |     | afterthechange.If |     | settofalseornullCreo |     |
| -------- | --------------- | --- | --- | ----------------- | --- | -------------------- | --- |
Parametricwilldelaythe repaint,allowing youto performseveraloperations
| beforeshowing   |          | changesonthescreen.          |     |         |         |       |                 |
| --------------- | -------- | ---------------------------- | --- | ------- | ------- | ----- | --------------- |
| Example:        | Creation |                              | of  | a Table | Listing | Datum | Points          |
| Thesamplecodein |          | thefilepfcDrawingExamples.vb |     |         |         |       | locatedat<creo_ |
vbapi_loadpoint>/vbapi_examplescreatesadrawing table thatlists
| thedatumpointsin |       | amodelshownin |          |     | adrawing | view. |     |
| ---------------- | ----- | ------------- | -------- | --- | -------- | ----- | --- |
| Drawing          | Table |               | Segments |     |          |       |     |
Drawingtables canbe constructedwith oneor moresegments.Eachsegment can
beindependently placed.The segmentsare specifiedby an integeridentifier
startingwith0.
MethodsandPropertyIntroduced:
• IpfcSelection.SelTableSegment
• IpfcTable.GetSegmentCount()
• IpfcTable.GetSegmentSheet()
• IpfcTable.MoveSegment()
• IpfcTable.GetInfo()
ThepropertyIpfcSelection.SelTableSegment returnsthevalue ofthe
segmentidentifier oftheselectedtable segment.Itreturns anullvalueif the
| selectiondoesnotcontain              |           |     | asegmentidentifier. |     |     |                    |     |
| ------------------------------------ | --------- | --- | ------------------- | --- | --- | ------------------ | --- |
| ThemethodIpfcTable.GetSegmentCount() |           |     |                     |     |     | returnsthenumberof |     |
| segmentsin                           | thetable. |     |                     |     |     |                    |     |
ThemethodIpfcTable.GetSegmentSheet() determinesthesheetnumber
| thatcontains | aspecifieddrawing |     |     | tablesegment. |     |     |     |
| ------------ | ----------------- | --- | --- | ------------- | --- | --- | --- |
ThemethodIpfcTable.MoveSegment() movesadrawing tablesegment to
anewlocation.Passthe coordinatesof thetarget positionin theformatx,y,z=0.
| 130 |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

Note
Setthe valueofthe parameterRepaint to trueto repaintthedrawing with the
| changes. | Setittofalseor | nullto | delaythe | repaint. |     |     |
| -------- | -------------- | ------ | -------- | -------- | --- | --- |
Toget informationaboutadrawing table passthevalueof thesegment identifier
themethodIpfcTable.GetInfo().
| asinputto |     |     |     | The methodreturnsthe |     | table |
| --------- | --- | --- | --- | -------------------- | --- | ----- |
informationincludingtherotation, rowand columninformation,andthe3D
outline.
| Repeat Regions |     |     |     |     |     |     |
| -------------- | --- | --- | --- | --- | --- | --- |
MethodsIntroduced:
• IpfcTable.IsCommentCell()
• IpfcTable.GetCellComponentModel()
• IpfcTable.GetCellReferenceModel()
• IpfcTable.GetCellTopModel()
• IpfcTableOwner.UpdateTables()
ThemethodsIpfcTable.IsCommentCell(),
IpfcTable.GetCellComponentModel(),
IpfcTable.GetCellReferenceModel(),
| IpfcTable.GetCellTopModel(),         |     |     | and |                 |     |         |
| ------------------------------------ | --- | --- | --- | --------------- | --- | ------- |
| IpfcTableOwner.UpdateTables()applyto |     |     |     | repeatregionsin |     | drawing |
tables.
| ThemethodIpfcTable.IsCommentCell()         |                   |     |     | tellsyouwhetheracellin |            | a      |
| ------------------------------------------ | ----------------- | --- | --- | ---------------------- | ---------- | ------ |
| repeatregion                               | containsacomment. |     |     |                        |            |        |
| ThemethodIpfcTable.GetCellComponentModel() |                   |     |     |                        | returnsthe | pathto |
theassemblycomponentmodelthatisbeingreferenced bya cellina repeatregion
ofadrawing table.It doesnotreturna validpathifthe cellattributeis setto NO
| DUPLICATEor                                | NODUPLICATE/LEVEL. |     |     |     |            |     |
| ------------------------------------------ | ------------------ | --- | --- | --- | ---------- | --- |
| ThemethodIpfcTable.GetCellReferenceModel() |                    |     |     |     | returnsthe |     |
referencecomponentthatisbeing referredto by acellin arepeatregionof a
| drawingtable,evenif | cellattributeis |     | settoNO | DUPLICATE | orNO |     |
| ------------------- | --------------- | --- | ------- | --------- | ---- | --- |
DUPLICATE/LEVEL.
ThemethodIpfcTable.GetCellTopModel()
|     |     |     |     | returnsthetop |     | modelthatis |
| --- | --- | --- | --- | ------------- | --- | ----------- |
beingreferredtoby acellin arepeatregion ofadrawing table,evenif cell
| attributeissetto | NO DUPLICATEor |     | NO DUPLICATE/LEVEL. |     |     |     |
| ---------------- | -------------- | --- | ------------------- | --- | --- | --- |
Drawings 131

Usethe methodIpfcTableOwner.UpdateTables()to updatetherepeat
regionsin allthe tablesto accountfor changesto themodel.It isequivalentto the
commandTable,RepeatRegion,Update.
| Detail              | Items |                         |     |              |
| ------------------- | ----- | ----------------------- | --- | ------------ |
| Themethodsdescribed |       | in thissectionoperateon |     | detailitems. |
IntheVB APIyoucan create,delete andmodifydetailitems,controltheir
display, andquerywhat detailitemsarepresent in thedrawing. Thetypesofdetail
itemsavailableare:
• DraftEntities—Containgraphicalitems createdin Creo Parametric.The items
areasfollows:
○ Arc
○ Ellipse
○ Line
○ Point
○ Polygon
○ Spline
• Notes—Textualannotations
| • SymbolDefinitions—Containedin      |     |     | thedrawing’s | symbol gallery. |
| ------------------------------------ | --- | --- | ------------ | --------------- |
| • SymbolInstances—Instancesofasymbol |     |     | placedin     | adrawing.       |
• DraftGroups—Groups ofdetailitems thatcontainnotes, symbol instances,
anddraftentities.
• OLEobjects—ObjectLinking andEmbedding(OLE)objectsembeddedin the
| CreoParametricdrawing |        | file. |     |     |
| --------------------- | ------ | ----- | --- | --- |
| Listing               | Detail | Items |     |     |
MethodsIntroduced:
• IpfcModelItemOwner.ListItems()
• IpfcDetailItemOwner.ListDetailItems()
• IpfcModelItemOwner.GetItemById()
• IpfcDetailItemOwner.CreateDetailItem()
ThemethodIpfcModelItemOwner.ListItems() returns alist ofdetail
itemsspecifiedby theparameterType orreturnsnullif nodetailitems ofthe
specifiedtypearefound.
| Thevalues | oftheparameter | Typefor | detailitemsare: |                                 |
| --------- | -------------- | ------- | --------------- | ------------------------------- |
| 132       |                |         |                 | Creo®ParametricVBAPIUser’sGuide |

• EpfcITEM_DTL_ENTITY—DetailEntity
• EpfcITEM_DTL_NOTE—DetailNote
• EpfcITEM_DTL_GROUP—DraftGroup
• EpfcITEM_DTL_SYM_DEFINITION—DetailSymbolDefinition
• EpfcITEM_DTL_SYM_INSTANCE—DetailSymbolInstance
| • EpfcITEM_DTL_OLE_OBJECT—Drawing |     |     | embeddedOLE | object |
| --------------------------------- | --- | --- | ----------- | ------ |
Ifthisparameter issetto null,thenall themodelitemsin thedrawing arelisted.
| Ifthemodelhasmultiplebodies,the |     | method |                     |     |
| ------------------------------- | --- | ------ | ------------------- | --- |
| IpfcModelItemOwner.ListItems()  |     |        | returnstheexception |     |
IpfcXToolkitMultibodyUnsupported.
| ThemethodIpfcDetailItemOwner.ListDetailItems() |     |     |     | also liststhe |
| ---------------------------------------------- | --- | --- | --- | ------------- |
detailitemsin themodel. Pass thetypeofthedetailitem andthe sheetnumber
| thatcontains | thespecifieddetailitems. |     |     |     |
| ------------ | ------------------------ | --- | --- | --- |
SettheinputparameterTypeto thetypeofdetailitem to belisted. Setitto nullto
returnallthe detailitems.The inputparameterSheetNumberdeterminesthe sheet
thatcontains thespecifieddetailitem.Pass nullto searchallthe sheets.This
| argumentis | ignorediftheparameterTypeis |     | settoEpfcDETAIL_SYM_ |     |
| ---------- | --------------------------- | --- | -------------------- | --- |
DEFINITION.
Themethodreturns asequenceof detailitemsand returnsanullif noitems
| matchingtheinputvalues |     | arefound. |     |     |
| ---------------------- | --- | --------- | --- | --- |
ThemethodIpfcModelItemOwner.GetItemById() returns adetailitem
basedon thetypeofthe detailitemand itsintegeridentifier.The methodreturnsa
| nullif adetailitem | withthe  | specifiedattributesisnot | found. |     |
| ------------------ | -------- | ------------------------ | ------ | --- |
| Creating           | a Detail | Item                     |        |     |
MethodsIntroduced:
• IpfcDetailItemOwner.CreateDetailItem()
• CCpfcDetailGroupInstructions.Create()
| ThemethodIpfcDetailItemOwner.CreateDetailItem() |     |     |     | createsa |
| ----------------------------------------------- | --- | --- | --- | -------- |
newdetailitembasedon theinstructiondataobjectthatdescribesthetypeand
contentofthe newdetailitem.The instructionsdataobjectis returnedbythe
| methodCCpfcDetailGroupInstructions.Create().The |     |             |     | method |
| ----------------------------------------------- | --- | ----------- | --- | ------ |
| returnsthenewlycreated                          |     | detailitem. |     |        |
Drawings 133

| Detail             | Entities |               |     |       |     |
| ------------------ | -------- | ------------- | --- | ----- | --- |
| Adetailentityinthe | VBAPI    | isrepresented |     | bythe |     |
interfaceIpfcDetailEntityItem.It isachild ofthe IpfcDetailItem
interface.
| TheinterfaceIpfcDetailEntityInstructions |                 |     |             | contains specific |     |
| ---------------------------------------- | --------------- | --- | ----------- | ----------------- | --- |
| informationusedto                        | describeadetail |     | entityitem. |                   |     |
Instructions
MethodsandPropertiesIntroduced:
• CCpfcDetailEntityInstructions.Create()
• IpfcDetailEntityInstructions.Geometry
• IpfcDetailEntityInstructions.IsConstruction
• IpfcDetailEntityInstructions.Color
• IpfcDetailEntityInstructions.FontName
• IpfcDetailEntityInstructions.Width
• IpfcDetailEntityInstructions.View
| ThemethodCCpfcDetailEntityInstructions.Create() |     |     |     |     | createsan |
| ----------------------------------------------- | --- | --- | --- | --- | --------- |
instructionsobjectthatdescribes howto construct adetailentity, forusein the
methodsIpfcDetailItemOwner.CreateDetailItem(),
IpfcDetailSymbolDefItem.CreateDetailItem(), and
IpfcDetailEntityItem.Modify().
Theinstructionsobjectiscreated basedon thecurvegeometryand thedrawing
viewassociatedwith theentity. Thecurvegeometry describesthetrajectory ofthe
detailentityinworld units. Thedrawing viewcan be amodelviewreturned bythe
methodIpfcModel2D.List2DViews()ora drawing sheetbackgroundview
| returnedbythe | method |     |     |     |     |
| ------------- | ------ | --- | --- | --- | --- |
IpfcSheetOwner.GetSheetBackgroundView().The backgroundview
| indicatesthattheentityisnot |                               | associatedwitha |     | particularmodelview. |     |
| --------------------------- | ----------------------------- | --------------- | --- | -------------------- | --- |
| Themethodreturns            | thecreatedinstructionsobject. |                 |     |                      |     |
Note
| Changesto | thevaluesof | aIpfcDetailEntityInstructions |     |     | objectdo |
| --------- | ----------- | ----------------------------- | --- | --- | -------- |
nottakeeffectuntil thatinstructionsobjectisusedto modifytheentityusing
IpfcDetailEntityItem.Modify().
| ThepropertyIpfcDetailEntityInstructions.Geometry |     |     |     |     | returnsthe |
| ------------------------------------------------ | --- | --- | --- | --- | ---------- |
geometryofthedetailentityitem.
| 134 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

ThepropertyIpfcDetailEntityInstructions.IsConstruction
returnsavaluethatspecifies whethertheentityisa constructionentity.
ThepropertyIpfcDetailEntityInstructions.Color returns thecolor
ofthedetailentityitem.
ThepropertyIpfcDetailEntityInstructions.FontName returnsthe
linestyleused todraw theentity. Themethodreturns anullvalueif thedefault
| linestyleis | used. |     |     |     |     |     |
| ----------- | ----- | --- | --- | --- | --- | --- |
ThepropertyIpfcDetailEntityInstructions.Width returns thevalue
ofthewidthof theentityline. Themethodreturns anullvalueif thedefaultline
widthisused.
| ThepropertyIpfcDetailEntityInstructions.View |     |     |     |     |     | returnsthe |
| -------------------------------------------- | --- | --- | --- | --- | --- | ---------- |
drawingview associatedwith theentity. Theview caneitherbe amodelviewor a
drawingsheetbackgroundview.
| Example:        | Create | a Draft                      | Line | with | Predefined | Color           |
| --------------- | ------ | ---------------------------- | ---- | ---- | ---------- | --------------- |
| Thesamplecodein |        | thefilepfcDrawingExamples.vb |      |      |            | locatedat<creo_ |
vbapi_loadpoint>/vbapi_examplesshowsautility thatcreatesadraft
| linein oneof | thecolorspredefinedin |             |     | Creo Parametric. |     |     |
| ------------ | --------------------- | ----------- | --- | ---------------- | --- | --- |
| Detail       | Entities              | Information |     |                  |     |     |
MethodandPropertyIntroduced:
• IpfcDetailEntityItem.GetInstructions()
• IpfcDetailEntityItem.SymbolDef
ThemethodIpfcDetailEntityItem.GetInstructions() returnsthe
| instructionsdataobjectthatisused          |     |     | toconstruct |     | the detailentityitem. |                   |
| ----------------------------------------- | --- | --- | ----------- | --- | --------------------- | ----------------- |
| ThepropertyIpfcDetailEntityItem.SymbolDef |     |     |             |     |                       | returns thesymbol |
definitionthatcontains theentity.This propertyreturnsa nullvalueifthe entityis
| notapart | ofasymbol | definition. |     |     |     |     |
| -------- | --------- | ----------- | --- | --- | --- | --- |
| Detail   | Entities  | Operations  |     |     |     |     |
MethodsIntroduced:
• IpfcDetailEntityItem.Draw()
• IpfcDetailEntityItem.Erase()
• IpfcDetailEntityItem.Modify()
ThemethodIpfcDetailEntityItem.Draw() temporarilydraws adetail
| entityitem, | sothatitisremoved |     | duringthenextdraftregeneration. |     |     |     |
| ----------- | ----------------- | --- | ------------------------------- | --- | --- | --- |
Drawings 135

| ThemethodIpfcDetailEntityItem.Erase() |     |     |     | undrawsa | detailentity |
| ------------------------------------- | --- | --- | --- | -------- | ------------ |
itemtemporarily,sothatitisredrawn duringthe nextdraftregeneration.
ThemethodIpfcDetailEntityItem.Modify() modifies thedefinitionof
| anentityitem | usingthe | specifiedinstructionsdataobject. |     |     |     |
| ------------ | -------- | -------------------------------- | --- | --- | --- |
| OLE          | Objects  |                                  |     |     |     |
Anobjectlinkingandembedding(OLE)objectisan externalfile,such asa
document,graphicsfile,orvideo filethatis createdusing anexternalapplication
andwhichcanbe insertedintoanotherapplication,such asCreo Parametric.You
cancreateandinsertsupported OLEobjectsintoatwo-dimensionalCreo
Parametricfile,suchas adrawing,report,formatfile, notebook,ordiagram. The
functionsdescribed inthis sectionenableyouto identifyandaccess OLEobjects
| embeddedin | drawings. |     |     |     |     |
| ---------- | --------- | --- | --- | --- | --- |
PropertiesIntroduced:
• IpfcDetailOLEObject.ApplicationType
• IpfcDetailOLEObject.Outline
• IpfcDetailOLEObject.Path
• IpfcDetailOLEObject.Sheet
ThepropertyIpfcDetailOLEObject.ApplicationType returnsthe type
| oftheOLE | objectasastring,for | example,MicrosoftWordDocument. |     |     |     |
| -------- | ------------------- | ------------------------------ | --- | --- | --- |
ThepropertyIpfcDetailOLEObject.Outline returnstheextentof the
| OLEobjectembeddedin |     | thedrawing. |     |     |     |
| ------------------- | --- | ----------- | --- | --- | --- |
ThepropertyIpfcDetailOLEObject.Path returnsthe pathto theexternal
| filefor eachOLEobject,if |     | itis linkedtoan | externalfile. |     |     |
| ------------------------ | --- | --------------- | ------------- | --- | --- |
ThepropertyIpfcDetailOLEObject.Sheet returnsthesheetnumberfor
theOLEobject.
| Detail        | Notes |                   |                |     |     |
| ------------- | ----- | ----------------- | -------------- | --- | --- |
| Adetailnotein | theVB | API isrepresented | bytheinterface |     |     |
IpfcDetailNoteItem.Itis achildof theIpfcDetailIteminterface .
TheinterfaceIpfcDetailNoteInstructions contains specificinformation
| thatdescribes | adetailnote. |     |     |     |     |
| ------------- | ------------ | --- | --- | --- | --- |
TheinterfaceIpfcDetailNoteInstructions andall themethodsunder
thisinterfacearedeprecated.
| 136 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

Instructions
MethodsandPropertiesIntroduced:
• CCpfcDetailNoteInstructions.Create()
• IpfcDetailNoteInstructions.TextLines
• IpfcDetailNoteInstructions.IsDisplayed
• IpfcDetailNoteInstructions.IsReadOnly
• IpfcDetailNoteInstructions.IsMirrored
• IpfcDetailNoteInstructions.Horizontal
• IpfcDetailNoteInstructions.Vertical
• IpfcDetailNoteInstructions.Color
• IpfcDetailNoteInstructions.Leader
• IpfcDetailNoteInstructions.TextAngle
| ThemethodCCpfcDetailNoteInstructions.Create() |     |     |     | createsadata |     |
| --------------------------------------------- | --- | --- | --- | ------------ | --- |
objectthatdescribes howadetailnoteitemshouldbe constructedwhenpassedto
themethodsIpfcDetailItemOwner.CreateDetailItem(),
| IpfcDetailSymbolDefItem.CreateDetailItem(), |     |     |     | or  |     |
| ------------------------------------------- | --- | --- | --- | --- | --- |
IpfcDetailNoteItem.Modify(). Theparameter inTextLinesspecifiesthe
sequenceoftextlinedataobjectsthatdescribethe contentsof thenote.
Note
Changesto thevaluesof aIpfcDetailNoteInstructionsobjectdo not
takeeffectuntil thatinstructionsobjectis usedto modifythenoteusing
IpfcDetailNoteItem.Modify().
| For creatinga                                         | detailnoteitem,themethod |     |     |                  |     |
| ----------------------------------------------------- | ------------------------ | --- | --- | ---------------- | --- |
| IpfcDetailItemOwner.CreateDetailItem()                |                          |     |     | is deprecated,as | it  |
| usesthedeprecatedinterfaceIpfcDetailNoteInstructions. |                          |     |     | Usethe           |     |
methodsIpfcDetailItemOwner.CreateFreeNote(),
IpfcDetailItemOwner.CreateOffsetNote(),
| IpfcDetailItemOwner.CreateOnItemNote(),           |                                      |                            |          | or         |     |
| ------------------------------------------------- | ------------------------------------ | -------------------------- | -------- | ---------- | --- |
| IpfcDetailItemOwner.CreateLeaderNote()            |                                      |                            |          | instead.   |     |
| Theproperty                                       | IpfcDetailNoteInstructions.TextLines |                            |          | returns    | the |
| descriptionof                                     | textlinecontents                     | in                         | thenote. |            |     |
| ThepropertyIpfcDetailNoteInstructions.IsDisplayed |                                      |                            |          | returnsa   |     |
| booleanindicatingifthe                            |                                      | noteis currentlydisplayed. |          |            |     |
| ThepropertyIpfcDetailNoteInstructions.IsReadOnly  |                                      |                            |          | determines |     |
| whetherthe                                        | notecanbe                            | editedbythe                | user.    |            |     |
| Drawings                                          |                                      |                            |          |            | 137 |

| ThepropertyIpfcDetailNoteInstructions.IsMirrored |                                 |     |     |          |            | determines |
| ------------------------------------------------ | ------------------------------- | --- | --- | -------- | ---------- | ---------- |
| whetherthe                                       | noteismirrored.                 |     |     |          |            |            |
| ThepropertyIpfcDetailNoteInstructions.Horizontal |                                 |     |     |          |            | returnsthe |
| valueofthe                                       | horizontaljustificationof       |     |     | thenote. |            |            |
| ThepropertyIpfcDetailNoteInstructions.Vertical   |                                 |     |     |          | returnsthe |            |
| valueofthe                                       | verticaljustificationofthenote. |     |     |          |            |            |
ThepropertyIpfcDetailNoteInstructions.Color returns thecolorof
thedetailnoteitem.The methodreturnsanullvalue torepresent thedefault
drawingcolor.
| ThepropertyIpfcDetailNoteInstructions.Leader    |     |     |                        |     | returnsthe |     |
| ----------------------------------------------- | --- | --- | ---------------------- | --- | ---------- | --- |
| locationsofthedetailnoteitem                    |     |     | andinformationaboutthe |     | leaders.   |     |
| ThepropertyIpfcDetailNoteInstructions.TextAngle |     |     |                        |     | returnsthe |     |
valueofthe angleofthe textusedin thenote.The methodreturns anullvalueif
| theangleis                                     | 0.0.       |                              |         |              |                 |      |
| ---------------------------------------------- | ---------- | ---------------------------- | ------- | ------------ | --------------- | ---- |
| Example:                                       | Create     | Drawing                      | Note    | at Specified | Location        | with |
| Leader                                         | to Surface | and                          | Surface | Name         |                 |      |
| Thesamplecodein                                |            | thefilepfcDrawingExamples.vb |         |              | locatedat<creo_ |      |
| vbapi_loadpoint>/vbapi_examplescreatesadrawing |            |                              |         |              | noteata         |      |
specifiedlocation,with aleaderattachedto asolidsurface,anddisplaysthename
ofthesurface.
| Detail | Notes | Information |     |     |     |     |
| ------ | ----- | ----------- | --- | --- | --- | --- |
MethodsandPropertyIntroduced:
• IpfcDetailNoteItem.GetInstructions()
• IpfcDetailNoteItem.SymbolDef
• IpfcDetailNoteItem.GetLineEnvelope()
• IpfcDetailNoteItem.GetModelReference()
| ThemethodIpfcDetailNoteItem.GetInstructions() |     |     |     |     | returnsan |     |
| --------------------------------------------- | --- | --- | --- | --- | --------- | --- |
instructionsdataobjectthatdescribeshowto construct thedetailnoteitem.This
| methodtakesaProBooleanargument, |     |     |     | GiveParametersAsNames,which |     |     |
| ------------------------------- | --- | --- | --- | --------------------------- | --- | --- |
determineswhethersymbolicrepresentationsof parametersand drawing
propertiesin the notetextshouldbe displayed, ortheactual textseen by theuser
shouldbedisplayed.
| 138 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | ------------------------------- | --- | --- |

Note
Creo Parametricdoesnotresolve andreplacesymboliccallouts fornotes
whicharenotdisplayed. Therefore, ifthenoteis notdisplayedorishidden in
alayer, thetextretrievedmaycontainsymboliccallouts,evenwhen
GiveParametersAsNames isfalse.
ThepropertyIpfcDetailNoteItem.SymbolDef returnsthe symbol
definitionthatcontains thenote.The methodreturnsa nullvalueifthe noteisnot
apartof asymbol definition.
ThemethodIpfcDetailNoteItem.GetLineEnvelope() determinesthe
screencoordinatesof theenvelopearoundthe detailnote.This envelopeisdefined
byfourpoints.The followingfigure illustrateshowthe pointorderis determined.
Theorderingof thepointsis maintainedevenifthe notesaremirroredor areatan
angle.
ThemethodIpfcDetailNoteItem.GetModelReference() returnsthe
modelreferenced bytheparameterizedtextin anote. Themodelis referenced
basedon thelinenumberandthe textindexwheretheparameterizedtextappears.
Details Notes Operations
MethodsIntroduced:
• IpfcDetailNoteItem.Draw()
• IpfcDetailNoteItem.Show()
• IpfcDetailNoteItem.Erase()
• IpfcDetailNoteItem.Remove()
• IpfcDetailNoteItem.KeepArrowTypeAsIs()
• IpfcDetailNoteItem.Modify()
ThemethodIpfcDetailNoteItem.Draw() temporarilydrawsa detailnote
item,sothatitisremoved duringthenextdraftregeneration.
ThemethodIpfcDetailNoteItem.Show() displaysthenoteitem,such
thatitis repaintedduringthe nextdraftregeneration.
Drawings 139

| ThemethodIpfcDetailNoteItem.Erase()  |                         |     |                           | undrawsa               | detailnoteitem |     |
| ------------------------------------ | ----------------------- | --- | ------------------------- | ---------------------- | -------------- | --- |
| temporarily,                         | sothatitisredrawnduring |     | thenextdraftregeneration. |                        |                |     |
| ThemethodIpfcDetailNoteItem.Remove() |                         |     |                           | undrawsadetailnoteitem |                |     |
permanently, sothatitisnotredrawn duringthenextdraftregeneration.
| ThemethodIpfcDetailNoteItem.KeepArrowTypeAsIs() |     |     |     |     |     | allowsyou |
| ----------------------------------------------- | --- | --- | --- | --- | --- | --------- |
tokeeparrow typeof theleadernoteasitis,afteranoteismodified.Youmust
| callthismethodbeforethe |     | methodIpfcDetailNoteItem.Modify()is |     |     |     |     |
| ----------------------- | --- | ----------------------------------- | --- | --- | --- | --- |
called.
ThemethodIpfcDetailNoteItem.Modify() modifiesthe definitionof an
existingdetailnoteitembasedon theinstructionsobjectthatdescribesthe new
detailnoteitem.
| Detail         | Groups |                     |              |     |     |     |
| -------------- | ------ | ------------------- | ------------ | --- | --- | --- |
| Adetailgroupin | theVB  | API isrepresentedby | theinterface |     |     |     |
IpfcDetailGroupItem.It isa childofthe IpfcDetailIteminterface.
TheinterfaceIpfcDetailGroupInstructions containsinformationused
| todescribe | adetailgroup | item. |     |     |     |     |
| ---------- | ------------ | ----- | --- | --- | --- | --- |
Instructions
MethodandPropertiesIntroduced:
• CCpfcDetailGroupInstructions.Create()
• IpfcDetailGroupInstructions.Name
• IpfcDetailGroupInstructions.Elements
• IpfcDetailGroupInstructions.IsDisplayed
| ThemethodCCpfcDetailGroupInstructions.Create() |     |     |     |     |     | createsan |
| ---------------------------------------------- | --- | --- | --- | --- | --- | --------- |
instructiondataobjectthatdescribeshowto construct adetailgroup forusein
pfcDetailItemOwner.CreateDetailItem()and
IpfcDetailGroupItem.Modify().
Note
| Changesto | thevaluesof | aIpfcDetailGroupInstructions |     |     |     | objectdo |
| --------- | ----------- | ---------------------------- | --- | --- | --- | -------- |
nottakeeffectuntil thatinstructionsobjectisusedto modifythegroup using
IpfcDetailGroupItem.Modify().
ThepropertyIpfcDetailGroupInstructions.Name returns thenameof
thedetailgroup.
| 140 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | ------------------------------- | --- | --- |

| ThepropertyIpfcDetailGroupInstructions.Elements |     |     |     |     |     | returnsthe |
| ----------------------------------------------- | --- | --- | --- | --- | --- | ---------- |
sequenceofthe detailitems(notes,groups andentities)contained in thegroup.
ThepropertyIpfcDetailGroupInstructions.IsDisplayed returns
| whetherthe | detailgroup | isdisplayedin |     | thedrawing. |     |     |
| ---------- | ----------- | ------------- | --- | ----------- | --- | --- |
| Detail     | Groups      | Information   |     |             |     |     |
MethodIntroduced:
• IpfcDetailGroupItem.GetInstructions()
ThemethodIpfcDetailGroupItem.GetInstructions() gets adata
objectthatdescribes howtoconstruct adetailgroup item.The methodreturns the
| dataobjectdescribing |        | thedetailgroup |     | item. |     |     |
| -------------------- | ------ | -------------- | --- | ----- | --- | --- |
| Detail               | Groups | Operations     |     |       |     |     |
MethodsIntroduced:
• IpfcDetailGroupItem.Draw()
• IpfcDetailGroupItem.Erase()
• IpfcDetailGroupItem.Modify()
| ThemethodIpfcDetailGroupItem.Draw() |          |               |     |                         | temporarily | drawsadetail |
| ----------------------------------- | -------- | ------------- | --- | ----------------------- | ----------- | ------------ |
| groupitem,so                        | thatitis | removedduring |     | thenextdraftgeneration. |             |              |
ThemethodIpfcDetailGroupItem.Erase() temporarilyundrawsa detail
| groupitem,so | thatitis | redrawnduring |     | thenextdraftgeneration. |     |     |
| ------------ | -------- | ------------- | --- | ----------------------- | --- | --- |
ThemethodIpfcDetailGroupItem.Modify() changesthe definitionof a
detailgroupitem basedonthe dataobjectthatdescribes howtoconstruct adetail
groupitem.
| Example:                                     | Create | New                          | Group | of Items |     |                 |
| -------------------------------------------- | ------ | ---------------------------- | ----- | -------- | --- | --------------- |
| Thesamplecodein                              |        | thefilepfcDrawingExamples.vb |       |          |     | locatedat<creo_ |
| vbapi_loadpoint>/vbapi_examplescreatesagroup |        |                              |       |          |     | from asetof     |
selecteddetailitems.
| Detail | Symbols |             |     |     |     |     |
| ------ | ------- | ----------- | --- | --- | --- | --- |
| Detail | Symbol  | Definitions |     |     |     |     |
Adetailsymboldefinitionin theVB APIis representedby theinterface
| IpfcDetailSymbolDefItem.It |     |     |     | isa childofthe | IpfcDetailItem |     |
| -------------------------- | --- | --- | --- | -------------- | -------------- | --- |
interface.
Drawings 141

TheinterfaceIpfcDetailSymbolDefInstructions containsinformation
thatdescribes asymbol definition.It canbe usedwhen creatingsymbol definition
| entitiesor | whileaccessing | existingsymbol | definitionentities. |     |
| ---------- | -------------- | -------------- | ------------------- | --- |
Instructions
MethodsandPropertiesIntroduced:
• CCpfcDetailSymbolDefInstructions.Create()
• IpfcDetailSymbolDefInstructions.SymbolHeight
• IpfcDetailSymbolDefInstructions.HasElbow
• IpfcDetailSymbolDefInstructions.IsTextAngleFixed
• IpfcDetailSymbolDefInstructions.ScaledHeight
• IpfcDetailSymbolDefInstructions.Attachments
• IpfcDetailSymbolDefInstructions.FullPath
• IpfcDetailSymbolDefInstructions.Reference
ThemethodCCpfcDetailSymbolDefInstructions.Create() creates
aninstructiondataobjectthatdescribeshowto createasymbol definitionbased
onthepathandname ofthesymbol definition. Theinstructionsobjectispassedto
themethodspfcDetailItemOwner.CreateDetailItem() and
pfcDetailSymbolDefItem.Modify().
Note
| Changesto | thevaluesof | aIpfcDetailSymbolDefInstructions |     |     |
| --------- | ----------- | -------------------------------- | --- | --- |
objectdo nottakeeffectuntilthatinstructionsobjectisused tomodifythe
| definitionusingthe |     | methodIpfcDetailSymbolDefItem.Modify(). |     |     |
| ------------------ | --- | --------------------------------------- | --- | --- |
ThepropertyIpfcDetailSymbolDefInstructions.SymbolHeight
returnsthevalue oftheheighttypeforthe symboldefinition. Thesymbol
| definitionheightoptions           |     | areasfollows: |        |     |
| --------------------------------- | --- | ------------- | ------ | --- |
| • EpfcSYMDEF_FIXED—Symbolheightis |     |               | fixed. |     |
EpfcSYMDEF_VARIABLE—Symbolheightisvariable.
•
• EpfcSYMDEF_RELATIVE_TO_TEXT—Symbolheightis determined
| relativetothe | textheight. |     |     |     |
| ------------- | ----------- | --- | --- | --- |
ThepropertyIpfcDetailSymbolDefInstructions.HasElbow
| determineswhetherthe |     | symboldefinitionincludesan |     | elbow. |
| -------------------- | --- | -------------------------- | --- | ------ |
Theproperty
IpfcDetailSymbolDefInstructions.IsTextAngleFixedreturns
| whetherthe | textoftheangle | isfixed. |     |                                 |
| ---------- | -------------- | -------- | --- | ------------------------------- |
| 142        |                |          |     | Creo®ParametricVBAPIUser’sGuide |

ThepropertyIpfcDetailSymbolDefInstructions.ScaledHeight
| returnstheheightofthe |     |     | symboldefinitionin |     | inches. |
| --------------------- | --- | --- | ------------------ | --- | ------- |
TheIpfcDetailSymbolDefInstructions.Attachments returns the
valueofthe sequenceofthe possibleinstanceattachmentpointsfor thesymbol
definition.
ThepropertyIpfcDetailSymbolDefInstructions.FullPath returns
| thevalueof | thecomplete |     | pathof | thesymbol | definitionfile. |
| ---------- | ----------- | --- | ------ | --------- | --------------- |
ThepropertyIpfcDetailSymbolDefInstructions.Reference returns
thetextreferenceinformationforthe symbol definition.Itreturns anullvalueif
thetextreferenceisnotused.The textreferenceidentifies thetextitem usedfora
symboldefinitionwhichhasaheighttypeofSYMDEF_TEXT_RELATED.
| Detail | Symbol | Definitions |     | Information |     |
| ------ | ------ | ----------- | --- | ----------- | --- |
MethodsIntroduced:
• IpfcDetailSymbolDefItem.ListDetailItems()
• IpfcDetailSymbolDefItem.GetInstructions()
ThemethodIpfcDetailSymbolDefItem.ListDetailItems()
lists the
detailitemsin thesymbol definitionbasedonthetypeof thedetailitem.
ThemethodIpfcDetailSymbolDefItem.GetInstructions() returns
aninstructiondataobjectthatdescribeshowto constructthe symbol definition.
| Detail | Symbol | Definitions |     | Operations |     |
| ------ | ------ | ----------- | --- | ---------- | --- |
MethodsIntroduced:
• IpfcDetailSymbolDefItem.CreateDetailItem()
• IpfcDetailSymbolDefItem.Modify()
ThemethodIpfcDetailSymbolDefItem.CreateDetailItem()
createsa detailitemin the symboldefinitionbasedonthe instructionsdataobject.
| Themethodreturns |     | thedetailitem |     | inthe | symboldefinition. |
| ---------------- | --- | ------------- | --- | ----- | ----------------- |
ThemethodIpfcDetailSymbolDefItem.Modify() modifies asymbol
definitionbasedon theinstructionsdataobjectthatcontainsinformationabout the
| modificationsto |        | bemade | tothe       | symbol definition. |     |
| --------------- | ------ | ------ | ----------- | ------------------ | --- |
| Retrieving      | Symbol |        | Definitions |                    |     |
MethodsIntroduced:
• IpfcDetailItemOwner.RetrieveSymbolDefItem()
Drawings 143

| FromCreo | 4.0F000onwards,the |     | method |     |     |     |
| -------- | ------------------ | --- | ------ | --- | --- | --- |
IpfcDetailItemOwner.RetrieveSymbolDefinition()hasbeen
| deprecated.Use                            |     | themethod |     |     |          |     |
| ----------------------------------------- | --- | --------- | --- | --- | -------- | --- |
| IpfcDetailItemOwner.RetrieveSymbolDefItem |     |           |     |     | instead. |     |
CreoParametricsymbols existin twodifferentareas: theuser-definedareaandthe
| systemsymbols |     | area. |     |     |     |     |
| ------------- | --- | ----- | --- | --- | --- | --- |
ThemethodIpfcDetailItemOwner.RetrieveSymbolDefItem()
retrievesasymbol definitionfromthe user-definedlocationdesignated bythe
configurationoption pro_symbol_dir. The symboldefinitionshouldhave
| beenpreviouslysavedto |     | afileusing | CreoParametric. |     |     |     |
| --------------------- | --- | ---------- | --------------- | --- | --- | --- |
ThemethodIpfcDetailItemOwner.RetrieveSymbolDefItem() also
retrievesasymbol definitionfromthe systemdirectory. Thesystem areacontains
symbolsprovided byCreo Parametricwiththe Detailmodule(such asthe
WeldingSymbolsLibrary).
| Theinputparametersofthis |     | methodare: |                 |     |     |     |
| ------------------------ | --- | ---------- | --------------- | --- | --- | --- |
| • FileName—Nameof        |     | thesymbol  | definitionfile. |     |     |     |
• Source—Sourceofthe symboldefinitionfile.The inputparameter Sourceis
| definedbythe                            |                                | enumeratedtype |     |                 |             |            |
| --------------------------------------- | ------------------------------ | -------------- | --- | --------------- | ----------- | ---------- |
| pfcDetail.DetailSymbolDefItemSource.The |                                |                |     |                 | validvalues | which      |
| aresupportedarelisted                   |                                | below:         |     |                 |             |            |
| ○                                       | DTLSYMDEF_SRC_SYSTEM—Specifies |                |     | thesystemsymbol |             | definition |
directory.
| ○   | DTLSYMDEF_SRC_PATH—Specifiesthe |     |             | absolutepathto |     | adirectory |
| --- | ------------------------------- | --- | ----------- | -------------- | --- | ---------- |
|     | containingthesymbol             |     | definition. |                |     |            |
• FilePath—Pathto thesymbol definitionfile.It isrelativeto thepathspecified
bythe optionpro_symbol_dir inthe configurationfile.Anullvalue
| indicatesthatthefunction |     |     | shouldsearchthecurrentdirectory. |     |     |     |
| ------------------------ | --- | --- | -------------------------------- | --- | --- | --- |
• Version—Numericalversion ofthe symboldefinitionfile.Anullvalue
| retrievesthe |     | latestversion. |     |     |     |     |
| ------------ | --- | -------------- | --- | --- | --- | --- |
• UpdateUnconditionally—Trueif Creoshouldupdateexisting instancesof this
symboldefinition, orfalseto quittheoperationifthe definitionexists inthe
model.
| Themethodreturns |     | theretrievedsymbol           | definition. |     |                 |     |
| ---------------- | --- | ---------------------------- | ----------- | --- | --------------- | --- |
| Example:Create   |     | SymbolDefinition             |             |     |                 |     |
| Thesamplecodein  |     | thefilepfcDrawingExamples.vb |             |     | locatedat<creo_ |     |
vbapi_loadpoint>/vbapi_examplescreatesasymbol definitionwhich
containsfourlineentitiesformingabox, anoteatthemiddle ofthebox, anda
freeattachment.
| 144 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | ------------------------------- | --- | --- |

Detail Symbol Instances
Adetailsymbolinstance in theVB APIis representedby theinterface
IpfcDetailSymbolInstItem.It isachild oftheIpfcDetailItem
interface.
TheinterfaceIpfcDetailSymbolInstInstructions contains
informationthatdescribesasymbol instance.It canbe usedwhen creatingsymbol
instancesandwhile accessingexisting groups.
Instructions
MethodsandPropertiesIntroduced:
• CCpfcDetailSymbolInstInstructions.Create()
• IpfcDetailSymbolInstInstructions.IsDisplayed
• IpfcDetailSymbolInstInstructions.Color
• IpfcDetailSymbolInstInstructions.SymbolDef
• IpfcDetailSymbolInstInstructions.AttachOnDefType
• IpfcDetailSymbolInstInstructions.DefAttachment
• IpfcDetailSymbolInstInstructions.InstAttachment
• IpfcDetailSymbolInstInstructions.Angle
• IpfcDetailSymbolInstInstructions.ScaledHeight
• IpfcDetailSymbolInstInstructions.TextValues
• IpfcDetailSymbolInstInstructions.CurrentTransform
• IpfcDetailSymbolInstInstructions.SetGroups()
ThemethodCCpfcDetailSymbolInstInstructions.Create()
createsa dataobjectthatcontainsinformationabout theplacementof asymbol
instance.
Note
Changesto thevaluesof aIpfcDetailSymbolInstInstructions
objectdo nottakeeffectuntilthatinstructionsobjectisused tomodifythe
instance usingIpfcDetailSymbolInstItem.Modify().
ThepropertyIpfcDetailSymbolInstInstructions.IsDisplayed
returnsavaluethatspecifies whethertheinstance ofthesymbol is displayed.
ThepropertyIpfcDetailSymbolInstInstructions.Color returnsthe
colorofthedetailsymbol instance.Anullvalue indicatesthatthedefaultdrawing
colorisused.
Drawings 145

ThepropertyIpfcDetailSymbolInstInstructions.SymbolDef
| returnsthesymbol |     | definitionusedforthe | instance. |     |     |     |
| ---------------- | --- | -------------------- | --------- | --- | --- | --- |
Theproperty
IpfcDetailSymbolInstInstructions.AttachOnDefTypereturns
theattachmenttypeof theinstance.The methodreturns anullvalueif the
attachmentrepresents afreeattachment. The attachmentoptions areas follows:
| • EpfcSYMDEFATTACH_FREE—Attachment             |            |         | onafree |     | point.       |     |
| ---------------------------------------------- | ---------- | ------- | ------- | --- | ------------ | --- |
| • EpfcSYMDEFATTACH_LEFT_LEADER—Attachmentvia   |            |         |         |     | aleaderonthe |     |
| leftside                                       | ofthe      | symbol. |         |     |              |     |
| • EpfcSYMDEFATTACH_RIGHT_LEADER—Attachment     |            |         |         |     | viaaleaderon | the |
| rightsideof                                    | thesymbol. |         |         |     |              |     |
| • EpfcSYMDEFATTACH_RADIAL_LEADER—Attachmentvia |            |         |         |     | aleaderata   |     |
radiallocation.
| • EpfcSYMDEFATTACH_ON_ITEM—Attachmenton |     |     |     | an  | itemin thesymbol |     |
| --------------------------------------- | --- | --- | --- | --- | ---------------- | --- |
definition.
| • EpfcSYMDEFATTACH_NORMAL_TO_ITEM—Attachmentnormalto |        |             |     |     |     | an item |
| ---------------------------------------------------- | ------ | ----------- | --- | --- | --- | ------- |
| inthe                                                | symbol | definition. |     |     |     |         |
ThepropertyIpfcDetailSymbolInstInstructions.DefAttachment
returnsthevalue thatrepresents thewayin whichthe instanceis attachedtothe
symboldefinition.
Theproperty
| IpfcDetailSymbolInstInstructions.InstAttachment |     |     |     |     | returnsthe |     |
| ----------------------------------------------- | --- | --- | --- | --- | ---------- | --- |
valueofthe attachmentof theinstance thatincludeslocationandleader
information.
| ThepropertyIpfcDetailSymbolInstInstructions.Angle |     |     |     |     | returnsthe |     |
| ------------------------------------------------- | --- | --- | --- | --- | ---------- | --- |
valueofthe angleatwhichtheinstance isplaced.The methodreturnsanullvalue
| ifthevalue | oftheangle | is0degrees. |     |     |     |     |
| ---------- | ---------- | ----------- | --- | --- | --- | --- |
ThepropertyIpfcDetailSymbolInstInstructions.ScaledHeight
returnstheheightofthe symbolinstance in theowner drawing ormodel
coordinates.This valueisconsistentwith theheightvalue shownfor asymbol
instancein thePropertiesdialogboxin theCreo ParametricUserInterface.
Note
The scaledheightobtainedusing theabovepropertyis partiallybasedonthe
| propertiesofthe                             |     | symboldefinitionassignedusing |     | theproperty |             |     |
| ------------------------------------------- | --- | ----------------------------- | --- | ----------- | ----------- | --- |
| IpfcDetailSymbolInstInstructions.SymbolDef. |     |                               |     |             | Changingthe |     |
symbol definitionmaychangethecalculated valuefor thescaled height.
| 146 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | ------------------------------- | --- | --- |

ThepropertyIpfcDetailSymbolInstInstructions.TextValues
returnsthesequence ofvarianttextvalues usedwhileplacingthe symbolinstance.
Theproperty
| IpfcDetailSymbolInstInstructions.CurrentTransform |     |          |                 | returns |
| ------------------------------------------------- | --- | -------- | --------------- | ------- |
| thecoordinatetransformationmatrixto               |     | placethe | symbolinstance. |         |
ThemethodIpfcDetailSymbolInstInstructions.SetGroups()
IpfcDetailSymbolGroupOption
| • EpfcDETAIL_SYMBOL_GROUP_INTERACTIVE—Symbol |     |     | groupsare |     |
| -------------------------------------------- | --- | --- | --------- | --- |
interactivelyselectedfor display.This isthedefault valuein theGRAPHICS
mode.
| • EpfcDETAIL_SYMBOL_GROUP_ALL—All |     |     | non-exclusivesymbol | groups |
| --------------------------------- | --- | --- | ------------------- | ------ |
areincludedfordisplay.
• EpfcDETAIL_SYMBOL_GROUP_NONE—Noneof thenon-exclusive symbol
| groupsare                               | includedfor | display. |                    |     |
| --------------------------------------- | ----------- | -------- | ------------------ | --- |
| • EpfcDETAIL_SYMBOL_GROUP_CUSTOM—Symbol |             |          | groups specifiedby |     |
theapplicationaredisplayed.
Refertothe sectionDetailSymbolGroupson page148formoreinformationon
detailsymbolgroups.
| Detail Symbol | Instances | Information |     |     |
| ------------- | --------- | ----------- | --- | --- |
MethodIntroduced:
• IpfcDetailSymbolInstItem.GetInstructions()
ThemethodIpfcDetailSymbolInstItem.GetInstructions()
returnsan instructionsdataobjectthatdescribeshowto construct asymbol
instance.Thismethod takesaProBooleanargument, GiveParametersAsNames,
whichdetermineswhethersymbolicrepresentationsof parametersand drawing
propertiesin the symbolinstance shouldbe displayed,or theactualtextseen by
| theusershouldbe | displayed. |            |     |     |
| --------------- | ---------- | ---------- | --- | --- |
| Detail Symbol   | Instances  | Operations |     |     |
MethodsIntroduced:
• IpfcDetailSymbolInstItem.Draw()
• IpfcDetailSymbolInstItem.Erase()
• IpfcDetailSymbolInstItem.Show()
• IpfcDetailSymbolInstItem.Remove()
• IpfcDetailSymbolInstItem.Modify()
| Drawings |     |     |     | 147 |
| -------- | --- | --- | --- | --- |

| ThemethodIpfcDetailSymbolInstItem.Draw()  |     |              |     |                              |     | drawsa symbol   |
| ----------------------------------------- | --- | ------------ | --- | ---------------------------- | --- | --------------- |
| instancetemporarily                       |     | tobe removed |     | onthe nextdraftregeneration. |     |                 |
| ThemethodIpfcDetailSymbolInstItem.Erase() |     |              |     |                              |     | undraws asymbol |
instancetemporarily fromthe displayto be redrawnon thenextdraftgeneration.
| ThemethodIpfcDetailSymbolInstItem.Show()   |                   |     |                        |     |     | displaysasymbol |
| ------------------------------------------ | ----------------- | --- | ---------------------- | --- | --- | --------------- |
| instanceto                                 | be repaintedonthe |     | nextdraftregeneration. |     |     |                 |
| ThemethodIpfcDetailSymbolInstItem.Remove() |                   |     |                        |     |     | deletesa symbol |
instancepermanently.
ThemethodIpfcDetailSymbolInstItem.Modify()
modifies asymbol
instancebasedon theinstructionsdataobjectthatcontainsinformationabout the
| modificationsto |     | bemade tothe                 | symbol | instance. |            |                 |
| --------------- | --- | ---------------------------- | ------ | --------- | ---------- | --------------- |
| Example:Create  |     | aFreeInstanceof              |        | Symbol    | Definition |                 |
| Thesamplecodein |     | thefilepfcDrawingExamples.vb |        |           |            | locatedat<creo_ |
vbapi_loadpoint>/vbapi_examplescreatesafree instanceof asymbol
definition.
Example:Create aFreeInstanceof aSymbol Definitionwith drawingunit
| heights,variabletext |     | andgroups                    |     |     |     |                 |
| -------------------- | --- | ---------------------------- | --- | --- | --- | --------------- |
| Thesamplecodein      |     | thefilepfcDrawingExamples.vb |     |     |     | locatedat<creo_ |
vbapi_loadpoint>/vbapi_examplescreatesafree instanceof asymbol
| definitionwith     |        | drawing unitheights, |       | variabletextandgroups. |              |     |
| ------------------ | ------ | -------------------- | ----- | ---------------------- | ------------ | --- |
| Detail             | Symbol | Groups               |       |                        |              |     |
| Adetailsymbolgroup |        | in theVB             | APIis | representedby          | theinterface |     |
IpfcDetailSymbolGroup.It isa childofthe IpfcObjectinterface.A
detailsymbolgroup isaccessible onlyas apart ofthecontents ofa detailsymbol
| definitionor                                  | instance. |     |     |     |     |          |
| --------------------------------------------- | --------- | --- | --- | --- | --- | -------- |
| TheinterfaceIpfcDetailSymbolGroupInstructions |           |     |     |     |     | contains |
informationthatdescribesasymbol group. Itcanbe usedwhen creatingnew
| symbolgroups, |     | orwhileaccessing | ormodifying |     | existinggroups. |     |
| ------------- | --- | ---------------- | ----------- | --- | --------------- | --- |
Instructions
MethodsandPropertiesIntroduced:
• CCpfcDetailSymbolGroupInstructions.Create()
• IpfcDetailSymbolGroupInstructions.Items
• IpfcDetailSymbolGroupInstructions.Name
| 148 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

ThemethodCCpfcDetailSymbolGroupInstructions.Create()
| createsthe | IpfcDetailSymbolGroupInstructions |     |     | dataobjectthat |     |
| ---------- | --------------------------------- | --- | --- | -------------- | --- |
storesthe nameofthe symbolgroup andthe listof detailitemsto be includedin
| thesymbol | group. |     |     |     |     |
| --------- | ------ | --- | --- | --- | --- |
Note
| Changesto | thevaluesof | theIpfcDetailSymbolGroupInstructions |     |     |     |
| --------- | ----------- | ------------------------------------ | --- | --- | --- |
dataobjectdonot takeeffectuntilthis objectisusedto modifytheinstance
| using                                              | themethodIpfcDetailSymbolGroup.Modify(). |             |        |     |            |
| -------------------------------------------------- | ---------------------------------------- | ----------- | ------ | --- | ---------- |
| ThepropertyIpfcDetailSymbolGroupInstructions.Items |                                          |             |        |     | returns    |
| thelistof                                          | detailitemsincludedin                    | thesymbol   | group. |     |            |
| ThepropertyIpfcDetailSymbolGroupInstructions.Name  |                                          |             |        |     | returnsthe |
| nameofthe                                          | symbolgroup.                             |             |        |     |            |
| Detail                                             | Symbol Group                             | Information |        |     |            |
MethodsIntroduced:
• IpfcDetailSymbolGroup.GetInstructions()
• IpfcDetailSymbolGroup.ParentGroup
• IpfcDetailSymbolGroup.ParentDefinition
• IpfcDetailSymbolGroup.ListChildren()
• IpfcDetailSymbolDefItem.ListSubgroups()
• IpfcDetailSymbolDefItem.IsSubgroupLevelExclusive()
• IpfcDetailSymbolDefItem.IsSubgroupLevelExclusive()
| ThemethodIpfcDetailSymbolGroup.GetInstructions() |     |     |     |     | returns the |
| ------------------------------------------------ | --- | --- | --- | --- | ----------- |
IpfcDetailSymbolGroupInstructionsdataobjectthatdescribeshow
| toconstruct                                     | asymbol group.       |        |                |            |            |
| ----------------------------------------------- | -------------------- | ------ | -------------- | ---------- | ---------- |
| ThemethodIpfcDetailSymbolGroup.ParentGroup      |                      |        |                | returnsthe | parent     |
| symbolgroup                                     | to whichagivensymbol |        | group belongs. |            |            |
| ThemethodIpfcDetailSymbolGroup.ParentDefinition |                      |        |                |            | returnsthe |
| symboldefinitionofa                             | givensymbolgroup.    |        |                |            |            |
| ThemethodIpfcDetailSymbolGroup.ListChildren()   |                      |        |                | liststhe   |            |
| subgroupsof                                     | agivensymbol         | group. |                |            |            |
ThemethodIpfcDetailSymbolDefItem.ListSubgroups()
lists the
subgroupsof agivensymbol group storedinthe symboldefinitionattheindicated
level.
| Drawings |     |     |     |     | 149 |
| -------- | --- | --- | --- | --- | --- |

Themethod
IpfcDetailSymbolDefItem.IsSubgroupLevelExclusive()
identifiesifthe subgroups ofagivensymbol group storedin thesymbol definition
atthe indicatedlevel areexclusiveorindependent.If groupsare exclusive,only
oneofthegroups atthislevel canbe activein themodelatanytime.If groupsare
independent,anynumberofgroups canbe active.
ThemethodIpfcDetailSymbolDefItem.ListSubgroups() lists the
symbolgroups includedin asymbol instance.The IpfcSymbolGroupFilter
argumentdeterminesthe typesof symbol groupsthatcan belisted. Ittakesthe
followingvalues:
• EpfcDTLSYMINST_ALL_GROUPS—Retrievesall groups inthe definition
ofthesymbol instance.
• EpfcDTLSYMINST_ACTIVE_GROUPS—Retrieves only thosegroups that
areactivelyshownin thesymbol instance.
• EpfcDTLSYMINST_INACTIVE_GROUPS—Retrievesonly thosegroups
thatarenotshownin thesymbol instance.
Detail Symbol Group Operations
MethodsIntroduced:
• IpfcDetailSymbolGroup.Delete()
• IpfcDetailSymbolGroup.Modify()
• IpfcDetailSymbolDefItem.CreateSubgroup()
• IpfcDetailSymbolDefItem.SetSubgroupLevelExclusive()
• IpfcDetailSymbolDefItem.SetSubgroupLevelIndependent()
ThemethodIpfcDetailSymbolGroup.Delete() deletes thespecified
symbolgroup from thesymbol definition.This methoddoesnotdeletethe entities
containedin thegroup.
ThemethodIpfcDetailSymbolGroup.Modify() modifies thespecified
symbolgroup basedonthe IpfcDetailSymbolGroupInstructionsdata
objectthatcontains informationaboutthe modificationsthatcanbe madeto the
symbolgroup.
ThemethodIpfcDetailSymbolDefItem.CreateSubgroup()creates a
newsubgroupin thesymbol definitionattheindicatedlevel belowtheparent
group.
Themethod
IpfcDetailSymbolDefItem.SetSubgroupLevelExclusive()
makesthesubgroups ofa symbolgroup exclusiveattheindicated level inthe
symboldefinition.
150 Creo®ParametricVBAPIUser’sGuide

Note
After yousetthesubgroups ofasymbol group asexclusive,only oneofthe
| groups attheindicated | levelcanbe |     | activein | themodelatanytime. |     |
| --------------------- | ---------- | --- | -------- | ------------------ | --- |
Themethod
IpfcDetailSymbolDefItem.SetSubgroupLevelIndependent()
makesthesubgroups ofa symbolgroup independentatthe indicatedlevel in the
symboldefinition.
Note
After yousetthesubgroups ofasymbol group asindependent,anynumberof
| groups attheindicated | levelcanbe |     | activein | themodelatanytime. |     |
| --------------------- | ---------- | --- | -------- | ------------------ | --- |
Detail Attachments
| Adetailattachmentin        | VB APIis | representedby |                    | the |     |
| -------------------------- | -------- | ------------- | ------------------ | --- | --- |
| interfaceIpfcAttachment.It |          | isusedfor     | thefollowingtasks: |     |     |
• Thewayin whicha drawingnoteora symbolinstance isplacedin adrawing.
• Thewayin whicha leaderon adrawing noteorsymbol instanceis attached.
MethodIntroduced:
• IpfcAttachment.GetType()
ThemethodIpfcAttachment.GetType()
returnsthe
IpfcAttachmentTypeobjectcontainingthetypesof detailattachments.The
| detailattachmenttypesareas |     | follows: |     |     |     |
| -------------------------- | --- | -------- | --- | --- | --- |
• EpfcATTACH_FREE—Theattachmentis ata freepointpossibly with respect
| toa givendrawing | view. |     |     |     |     |
| ---------------- | ----- | --- | --- | --- | --- |
• EpfcATTACH_PARAMETRIC—The attachmentis to apointon asurfaceor
| anedgeof asolid. |     |     |     |     |     |
| ---------------- | --- | --- | --- | --- | --- |
• EpfcATTACH_OFFSET—The attachmentis offsetto anotherdrawing view,
| toa modelitem,orto                | a3D | modelannotation. |     |                |             |
| --------------------------------- | --- | ---------------- | --- | -------------- | ----------- |
| • EpfcATTACH_TYPE_UNSUPPORTED—The |     |                  |     | attachmentisto | an itemthat |
cannotberepresented in PFCatthe currenttime.However,you canstill
| retrievethe locationoftheattachment. |     |     |     |     |     |
| ------------------------------------ | --- | --- | --- | --- | --- |
Drawings 151

Free Attachment
TheEpfcATTACH_FREE detailattachmenttypeisrepresented bytheinterface
IpfcFreeAttachment.Itis achildof theIpfcAttachmentinterface .
PropertiesIntroduced:
• IpfcFreeAttachment.AttachmentPoint
• IpfcFreeAttachment.View
ThepropertyIpfcFreeAttachment.AttachmentPoint returns the
attachmentpoint.This locationisin screencoordinatesfor drawingitems, symbol
instancesandsurface finishes onflat-to-screenannotationplanes,andin model
coordinatesfor symbols andsurfacefinishes on 3Dmodelannotationplanes.
ThemethodIpfcFreeAttachment.View returnsthedrawing view towhich
theattachmentisrelated.The attachmentpoint isrelativeto thedrawing view, that
istheattachmentpointmoveswhen thedrawing view ismoved.This method
returnsaNULLvalue,if thedetailattachmentisnotrelated toa drawing view, but
isplacedatthespecifiedlocationin thedrawing sheet,or ifthe attachmentis
offsetto amodelitem orto a3Dmodelannotation.
Parametric Attachment
TheEpfcATTACH_PARAMETRIC detailattachmenttypeisrepresented bythe
interfaceIpfcParametricAttachment.It isachild ofthe
IpfcAttachmentinterface.
PropertyIntroduced:
• IpfcParametricAttachment.AttachedGeometry
ThepropertyIpfcParametricAttachment.AttachedGeometry returns
theIpfcSelection objectrepresentingtheitemto whichthe detailattachment
isattached.Thisincludes thedrawing view inwhichthe attachmentis made.
Offset Attachment
TheEpfcATTACH_OFFSET detailattachmenttypeisrepresented bythe
interfaceIpfcOffsetAttachment.It isachild oftheIpfcAttachment
interface.
PropertiesIntroduced:
• IpfcOffsetAttachment.AttachedGeometry
• IpfcOffsetAttachment.AttachmentPoint
152 Creo®ParametricVBAPIUser’sGuide

ThepropertyIpfcOffsetAttachment.AttachedGeometry returnsthe
IpfcSelectionobjectrepresentingthe itemto whichthedetailattachmentis
attached.This includesthe drawing viewwheretheattachmentismade, ifthe
offsetreferenceisin amodel.
ThepropertyIpfcOffsetAttachment.AttachmentPoint returnsthe
attachmentpoint.This locationisin screencoordinatesfor drawingitems, symbol
instancesandsurface finishes onflat-to-screenannotationplanes,andin model
coordinatesfor symbols andsurfacefinishes on 3Dmodelannotationplanes.The
distancefrom theattachmentpointto thelocationof theitemto whichthe detail
attachmentis attachedissavedasthe offsetdistance.
Unsupported Attachment
TheEpfcATTACH_TYPE_UNSUPPORTEDdetail attachmenttypeis represented
bytheinterfaceIpfcUnsupportedAttachment. Itis achild ofthe
IpfcAttachmentinterface.
PropertyIntroduced:
• IpfcUnsupportedAttachment.AttachmentPoint
ThepropertyIpfcUnsupportedAttachment.AttachmentPoint returns
theattachmentpoint.Thislocationis inscreen coordinatesfor drawing items,
symbolinstances andsurface finisheson flat-to-screenannotationplanes, andin
modelcoordinatesfor symbolsand surfacefinishes on3D modelannotation
planes.
Drawings 153

11
Solid
GettingaSolidObject..............................................................................................155
SolidInformation.....................................................................................................155
SolidOperations......................................................................................................155
SolidUnits..............................................................................................................158
MassProperties......................................................................................................164
Annotations............................................................................................................166
CrossSections........................................................................................................166
Materials.................................................................................................................167
Mostofthe objectsandmethodsin theVB APIareused withsolidmodels(parts
andassemblies).Becausesolidobjectsinheritfrom theinterfaceIpfcModel,
youcanuseanyofthe IpfcModelmethods onanyIpfcSolid, IpfcPart,
orIpfcAssemblyobject.
154 Creo®ParametricVBAPIUser’sGuide

| Getting | a Solid | Object |     |     |
| ------- | ------- | ------ | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcBaseSession.CreatePart()
• IpfcBaseSession.CreateAssembly()
• IpfcComponentPath.Root
• IpfcComponentPath.Leaf
• IpfcMFG.GetSolid()
ThemethodsIpfcBaseSession.CreatePart() and
IpfcBaseSession.CreateAssembly() createnewsolidmodelswith the
namesyouspecify.
| ThepropertiesIpfcComponentPath.Root |     |     | and                       |     |
| ----------------------------------- | --- | --- | ------------------------- | --- |
| IpfcComponentPath.Leafspecify       |     |     | thesolidobjectsthatmakeup | the |
componentpathofan assembly componentmodel. You canget acomponentpath
| objectfrom | anycomponentthathasbeeninteractivelyselected. |     |     |     |
| ---------- | --------------------------------------------- | --- | --- | --- |
ThemethodIpfcMFG.GetSolid() retrievesthe storagesolidin whichthe
manufacturingmodel’s featuresareplaced.Inorder to createaUDFgroup in the
manufacturingmodel,callthe methodIpfcSolid.CreateUDFGroup() on
| thestorage        | solid. |     |     |     |
| ----------------- | ------ | --- | --- | --- |
| Solid Information |        |     |     |     |
PropertiesIntroduced:
• IpfcSolid.RelativeAccuracy
• IpfcSolid.AbsoluteAccuracy
You cansettherelativeandabsolute accuracyofanysolidmodelusing these
methods.Relativeaccuracyisrelativeto thesize ofthesolid. For example,a
relativeaccuracyof .01specifiesthatthe solidmustbe accurateto within1/100 of
itssize. Absoluteaccuracyismeasuredin absoluteunits(inches,centimeters, and
soon).
Note
For achangein accuracyto takeeffect,youmustregeneratethemodel.
| Solid Operations |     |     |     |     |
| ---------------- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
| Solid |     |     |     | 155 |
| ----- | --- | --- | --- | --- |

• IpfcSolid.Regenerate()
• CCpfcRegenInstructions.Create()
• IpfcRegenInstructions.AllowFixUI
• IpfcRegenInstructions.ForceRegen
• IpfcRegenInstructions.FromFeat
• IpfcRegenInstructions.RefreshModelTree
• IpfcRegenInstructions.ResumeExcludedComponents
• IpfcRegenInstructions.UpdateAssemblyOnly
• IpfcRegenInstructions.UpdateInstances
• IpfcRegenInstructions.ResolveModeRegen
• IpfcSolid.GeomOutline
• IpfcSolid.EvalOutline()
• IpfcSolid.IsSkeleton
• IpfcSolid.ListGroups()
• IpfcSolid.GetSurfaceSolidBody()
• IpfcSolid.GetEdgeSolidBody()
ThemethodIpfcSolid.Regenerate()
|                |                      | causesthesolidmodelto |     | regenerate |
| -------------- | -------------------- | --------------------- | --- | ---------- |
| accordingtothe | instructionsprovided | in theformof          | the |            |
IpfcRegenInstructions object.Passing anullvaluefor theinstructions
| argumentcausesan | automaticregeneration. |     |     |     |
| ---------------- | ---------------------- | --- | --- | --- |
Pro/ENGINEERWildfire5.0 introducestheNo-Resolvemode,whereinif a
modelandfeatureregeneration fails,failedfeaturesandchildren offailedfeatures
arecreatedandregeneration ofotherfeaturescontinues. However,VB API does
| notsupportregeneration       | inthis mode.Themethod |              |     |     |
| ---------------------------- | --------------------- | ------------ | --- | --- |
| IpfcSolid.Regenerate()throws |                       | an exception |     |     |
IpfcXToolkitBadContext,ifCreo Parametricisrunning in theNo-Resolve
mode.Tocontinuewith thePro/ENGINEERWildfire4.0behavior inthe Resolve
| mode,setthe    | configurationoptionregen_failure_handling |     |     | to  |
| -------------- | ----------------------------------------- | --- | --- | --- |
| resolve_modein | theCreoParametricsession.                 |     |     |     |
Note
Settingthe configurationoption to switchto Resolvemodeensuresthe old
behavioras longas youdonot retrievethe modelssavedundertheNo-
Resolvemode.Toconsistentlypreservetheoldbehavior,useResolvemode
| from thebeginningandthroughoutyour |     | CreoParametricsession. |                                 |     |
| ---------------------------------- | --- | ---------------------- | ------------------------------- | --- |
| 156                                |     |                        | Creo®ParametricVBAPIUser’sGuide |     |

TheIpfcRegenInstructions objectcontains thefollowinginput
parameters:
• AllowFixUI—Determineswhetherornotto activatethe FixModeluser
interface,ifthereis anerror.
Usethe propertyIpfcRegenInstructions.AllowFixUI tomodify
thisparameter.
• ForceRegen—CreoParametric
Usethe propertyIpfcRegenInstructions.ForceRegen tomodify
thisparameter.
• FromFeat—Notcurrentlyused.This parameteris reservedfor futureuse.
Usethe propertyIpfcRegenInstructions.FromFeat to modifythis
parameter.
• RefreshModelTree—Creo ParametricModelTree
Usethe propertyIpfcRegenInstructions.RefreshModelTree to
modifythisparameter.
• ResumeExcludedComponents—CreoParametric
Usethe methodproperty
IpfcRegenInstructions.ResumeExcludedComponents to
modifythisparameter.
• UpdateAssemblyOnly—Updatestheplacementsof anassembly andallitssub-
assemblies,andregeneratestheassembly featuresandintersectedparts.Ifthe
affectedassembly isretrievedasa simplifiedrepresentation,thenthe locations
ofthecomponentsareupdated.If thisattributeis false,thecomponent
locationsarenotupdated,evenifthe simplifiedrepresentationisretrieved.By
default,itisfalse.
Usethe propertyIpfcRegenInstructions.UpdateAssemblyOnly
tomodifythis parameter.
• UpdateInstances—Updatestheinstances ofthesolidmodelin memory. This
mayslowdowntheregeneration process.Bydefault,this attributeisfalse.
Usethe propertyIpfcRegenInstructions.UpdateInstances to
modifythisparameter.
• ResolveModeRegen—Allowsregeneration ofa solidin resolve mode.The
regenerationbehavior iscontrolledbytemporarily overridingthedefault
settings.
Usethe propertyIpfcRegenInstructions.ResolveModeRegen to
modifythisparameter.Bydefault,itisfalse. Ifyouwant tosetyour Creo
Solid 157

Parametricapplicationin resolve mode,youneedtosetthe valueofthis
true.
parameterto
Note
|     | However,resolve                    |             | modewillbedeprecatedin |                          | afuturereleaseof | Creo |
| --- | ---------------------------------- | ----------- | ---------------------- | ------------------------ | ---------------- | ---- |
|     | Parametric.Hence,itisrecommendedto |             |                        | runtheapplicationwithout |                  |      |
|     | calling                            | theproperty |                        |                          |                  |      |
IpfcRegenInstructions.ResolveModeRegen.
ThepropertyIpfcSolid.GeomOutline returns thethree-dimensional
| boundingboxfor |     | thespecifiedsolid. |     |     |     |     |
| -------------- | --- | ------------------ | --- | --- | --- | --- |
Note
Donot useIpfcSolid.GeomOutline tocalculatethe outlineofa solidas
thedimensions oftheboundary boxcouldbe slightlybigger thanthe outline
|     | dimensionsof | thegeometry.       |     | UseIpfcSolid.EvalOutline() |     | to  |
| --- | ------------ | ------------------ | --- | -------------------------- | --- | --- |
|     | computean    | accurateoutlineofa |     | solid.                     |     |     |
ThemethodIpfcSolid.EvalOutline() alsoreturnsa three-dimensional
boundingbox, butyoucanspecify thecoordinatesystemused tocomputethe
extentsofthesolidobject.
ThepropertyIpfcSolid.IsSkeleton determineswhetherthepartmodelis
askeletonor aconceptmodel. Itreturns atrue valueif themodelisaskeleton,
| elseitreturnsa |     | false. |     |     |     |     |
| -------------- | --- | ------ | --- | --- | --- | --- |
ThemethodIpfcSolid.ListGroups() returnsthe listof groups(including
| UDFs)in                                  | thesolid. |     |     |     |            |        |
| ---------------------------------------- | --------- | --- | --- | --- | ---------- | ------ |
| ThemethodIpfcSolid.GetSurfaceSolidBody() |           |     |     |     | returnsthe | bodyto |
whichthesurfacebelongs.
ThemethodIpfcSolid.GetEdgeSolidBody() returns thebodyto which
theedgebelongs.
| Solid | Units |     |     |     |     |     |
| ----- | ----- | --- | --- | --- | --- | --- |
Eachmodelhasa basicsystemofunitsto ensureall materialpropertiesofthat
modelareconsistently measuredanddefined.Allmodelsaredefinedon thebasis
| ofthesystem |     | ofunits. | Apartcan | haveonly onesystem | ofunit. |     |
| ----------- | --- | -------- | -------- | ------------------ | ------- | --- |
Thefollowingtypesofquantitiesgovernthedefinitionofunitsof measurement:
| 158 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

• BasicQuantities—Thebasicunitsanddimensions ofthesystem ofunits. For
| example,consider      | theCentimeterGramSecond(CGS)system |           |     | ofunit. |
| --------------------- | ---------------------------------- | --------- | --- | ------- |
| Thebasicquantitiesfor | thissystemof                       | unitsare: |     |         |
Length—cm
○
○ Mass—g
○ Force—dyne
Time—sec
○
○ Temperature—K
• DerivedQuantities—Thederivedunitsarethosethatarederivedfromthe
| basicquantities.For | example,consider | theCentimeterGramSecond |     |     |
| ------------------- | ---------------- | ----------------------- | --- | --- |
(CGS)system ofunit. Thederivedquantitiesforthis systemof unitare as
follows:
Area—cm^2
○
○ Volume—cm^3
○ Velocity—cm/sec
IntheVB API,individualunitsin themodelarerepresentedby theinterface
pfcUnits.Unit.
| Types of Unit     | Systems      |          |     |     |
| ----------------- | ------------ | -------- | --- | --- |
| Thetypesofsystems | ofunitsareas | follows: |     |     |
• Pre-definedsystem ofunits—This systemof unitis providedby default.
• Custom-definedsystem ofunits—This systemof unitis definedbythe user
onlyif themodeldoesnotcontain standardmetricor nonmetricunits, orif the
materialfilecontainsunitsthatcannotbe derivedfromthe predefinedsystem
ofunitsorboth.
| InCreo Parametric,the | systemofunitsarecategorized |     | asfollows: |     |
| --------------------- | --------------------------- | --- | ---------- | --- |
• MassLengthTime(MLT)—Thefollowingsystems ofunitsbelong tothis
category:
| ○ CGS—Centimeter | GramSecond     |     |     |     |
| ---------------- | -------------- | --- | --- | --- |
| ○ MKS—Meter      | KilogramSecond |     |     |     |
○ mmKS—millimeterKilogramSecond
• ForceLengthTime(FLT)—Thefollowingsystemsof unitsbelongto this
category:
○ Creo ParametricDefault—InchlbmSecond. Thisis thedefaultsystem
| followedby | Creo Parametric. |     |     |     |
| ---------- | ---------------- | --- | --- | --- |
| Solid      |                  |     |     | 159 |

| ○ FPS—FootPound | Second |     |     |     |
| --------------- | ------ | --- | --- | --- |
| ○ IPS—InchPound | Second |     |     |     |
○ mmNS—MillimeterNewtonSecond
IntheVB API,thesystem ofunitsfollowedby themodelis representedby the
interfaceIpfcUnitSystem.
| Accessing | Individual | Units |     |     |
| --------- | ---------- | ----- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcSolid.ListUnits()
• IpfcSolid.GetUnit()
• IpfcUnit.Name
• IpfcUnit.Expression
• IpfcUnit.Type
• IpfcUnit.IsStandard
• IpfcUnit.ReferenceUnit
• IpfcUnit.ConversionFactor
• IpfcUnitConversionFactor.Offset
• IpfcUnitConversionFactor.Scale
ThemethodIpfcSolid.ListUnits() returns thelist ofunitsavailableto
thespecifiedmodel.
ThemethodIpfcSolid.GetUnit() retrieves theunit, basedonitsname or
expressionfor thespecifiedmodelinthe formof theIpfcUnitobject.
| ThepropertyIpfcUnit.Name |     | returnsthe | nameofthe | unit. |
| ------------------------ | --- | ---------- | --------- | ----- |
ThepropertyIpfcUnit.Expression returns auser-friendlyunitdescription
inthe formof thename(forexample,ksi)for ordinaryunitsandthe expression
| (forexample,N/m^3) | forsystem-generated |     | units. |     |
| ------------------ | ------------------- | --- | ------ | --- |
ThepropertyIpfcUnitSystem.Type
returns thetypeofquantityrepresented
bytheunitin termsof theIpfcUnitTypeobject.Thetypesof unitsare as
follows:
| • EpfcUNIT_LENGTH—Specifies |     | lengthmeasurementunits. |     |     |
| --------------------------- | --- | ----------------------- | --- | --- |
EpfcUNIT_MASS—Specifiesmass
| •   |     |     | measurementunits. |     |
| --- | --- | --- | ----------------- | --- |
• EpfcUNIT_FORCE—Specifiesforcemeasurementunits.
| • EpfcUNIT_TIME—Specifiestime |     |     | measurementunits. |     |
| ----------------------------- | --- | --- | ----------------- | --- |
EpfcUNIT_TEMPERATURE—Specifiestemperaturemeasurementunits.
•
• EpfcUNIT_ANGLE—Specifiesanglemeasurementunits.
| 160 |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

ThepropertyIpfcUnit.IsStandard identifies whethertheunitissystem-
defined(ifthe propertyIsStandardis setto true)or user-defined(if theproperty
| IsStandardis |     | settofalse). |     |     |     |     |     |     |     |     |
| ------------ | --- | ------------ | --- | --- | --- | --- | --- | --- | --- | --- |
ThepropertyIpfcUnit.ReferenceUnit returnsareferenceunit(oneof the
| availablesystem |     | units)in |     | termsoftheIpfcUnit |     |     |     | object. |     |     |
| --------------- | --- | -------- | --- | ------------------ | --- | --- | --- | ------- | --- | --- |
ThepropertyIpfcUnit.ConversionFactor identifiestherelationofthe
| unitto     | itsreferenceunitinterms |                                    |     |     | oftheIpfcUnitConversionFactor |     |     |     |     |     |
| ---------- | ----------------------- | ---------------------------------- | --- | --- | ----------------------------- | --- | --- | --- | --- | --- |
| object.The |                         | unitconversionfactorsareasfollows: |     |     |                               |     |     |     |     |     |
• Offset—Specifiestheoffset valueappliedto the valuesin thereference
unit.
• Scale—Specifiesthescale appliedto thevalues inthe referenceunitto get
|     | thevaluein | the        | actualunit. |     |         |     |         |             |     |      |
| --- | ---------- | ---------- | ----------- | --- | ------- | --- | ------- | ----------- | --- | ---- |
|     | Example    | - Consider |             | the | formula | to  | convert | temperature |     | from |
Centigrade
|     | to Fahrenheit |        |     |     |     |     |     |     |     |     |
| --- | ------------- | ------ | --- | --- | --- | --- | --- | --- | --- | --- |
|     | F = a         | + (C * | b)  |     |     |     |     |     |     |     |
where
|     | F is the | temperature |            | in  | Fahrenheit |            |     |           |     |     |
| --- | -------- | ----------- | ---------- | --- | ---------- | ---------- | --- | --------- | --- | --- |
|     | C is the | temperature |            | in  | Centigrade |            |     |           |     |     |
|     | a = 32   | (constant   | signifying |     |            | the offset |     | value)    |     |     |
|     | b = 9/5  | (ratio      | signifying |     | the        | scale      | of  | the unit) |     |     |
Note
|                                | Creo                                      | Parametricscalesthelengthdimensions |           |         |                            |     |            | ofthemodelusing     |              | the |
| ------------------------------ | ----------------------------------------- | ----------------------------------- | --------- | ------- | -------------------------- | --- | ---------- | ------------------- | ------------ | --- |
|                                | factors                                   | listedabove.Ifthe                   |           |         | scaleismodified,themodelis |     |            |                     | regenerated. |     |
|                                | Whenyouscale                              |                                     | themodel, |         | themodelunitsare           |     |            | notchanged.Imported |              |     |
|                                | geometrycannotbe                          |                                     |           | scaled. |                            |     |            |                     |              |     |
| Usethe                         | propertiesIpfcUnitConversionFactor.Offset |                                     |           |         |                            |     |            |                     | and          |     |
| IpfcUnitConversionFactor.Scale |                                           |                                     |           |         |                            |     | toretrieve | theunitconversion   |              |     |
| factorslisted                  |                                           | above.                              |           |         |                            |     |            |                     |              |     |
| Modifying                      |                                           | Individual                          |           |         | Units                      |     |            |                     |              |     |
MethodsandPropertiesIntroduced:
• IpfcUnit.Modify()
• IpfcUnit.Delete()
| ThemethodIpfcUnit.Modify()         |     |                                   |     |     |     | modifies | thedefinitionofa |     | unitby |     |
| ---------------------------------- | --- | --------------------------------- | --- | --- | --- | -------- | ---------------- | --- | ------ | --- |
| applyinga                          |     | newconversionfactorspecifiedbythe |     |     |     |          |                  |     |        |     |
| IpfcUnitConversionFactorobjectanda |     |                                   |     |     |     |          | referenceunit.   |     |        |     |
Solid 161

| ThemethodIpfcUnit.Delete() |     | deletesthe | unit. |     |
| -------------------------- | --- | ---------- | ----- | --- |
Note
| Youcandelete | onlycustomunitsandnotstandardunits. |     |     |     |
| ------------ | ----------------------------------- | --- | --- | --- |
| Creating     | a New Unit                          |     |     |     |
MethodsIntroduced:
• IpfcSolid.CreateCustomUnit()
• CCpfcUnitConversionFactor.Create()
ThemethodIpfcSolid.CreateCustomUnit() createsa customunitbased
| onthespecifiedname,the | conversionfactorgivenby |     | the |     |
| ---------------------- | ----------------------- | --- | --- | --- |
IpfcUnitConversionFactorobject,andareferenceunit.
| ThemethodCCpfcUnitConversionFactor.Create() |     |     | createsthe |     |
| ------------------------------------------- | --- | --- | ---------- | --- |
IpfcUnitConversionFactorobjectcontainingtheunitconversionfactors.
| Accessing | Systems | of Units |     |     |
| --------- | ------- | -------- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcSolid.ListUnitSystems()
• IpfcSolid.GetPrincipalUnits()
• IpfcUnitSystem.GetUnit()
• IpfcUnitSystem.Name
• IpfcUnitSystem.Type
• IpfcUnitSystem.IsStandard
| ThemethodIpfcSolid.ListUnitSystems()   |                    |     | returnsthelist   | ofunit |
| -------------------------------------- | ------------------ | --- | ---------------- | ------ |
| systemsavailableto                     | thespecifiedmodel. |     |                  |        |
| ThemethodIpfcSolid.GetPrincipalUnits() |                    |     | returnsthesystem | of     |
oftheIpfcUnitSystem
| unitsassignedto | thespecifiedmodelin | theform |     |     |
| --------------- | ------------------- | ------- | --- | --- |
object.
ThemethodIpfcUnitSystem.GetUnit() retrieves theunitofa particular
| typeused bytheunitsystem. |     |     |     |     |
| ------------------------- | --- | --- | --- | --- |
ThepropertyIpfcUnitSystem.Name returns thenameof theunitsystem.
ThepropertyIpfcUnitSystem.Type returns thetypeoftheunitsystem in
theformof theIpfcUnitSystemType object.The typesofunitsystems areas
follows:
| 162 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

• EpfcUNIT_SYSTEM_MASS_LENGTH_TIME—Specifies theMassLength
| Time(MLT) | unitsystem. |     |
| --------- | ----------- | --- |
• EpfcUNIT_SYSTEM_FORCE_LENGTH_TIME—Specifiesthe ForceLength
Time(FLT)unitsystem.
Formoreinformationon theseunitsystemslisted above,refer tothe sectionTypes
| ofUnit Systemson | page159. |     |
| ---------------- | -------- | --- |
ThepropertyIpfcUnitSystem.IsStandard identifieswhethertheunit
systemissystem-defined(if the propertyIsStandardis settotrue)or user-defined
| (ifthe propertyIsStandardis |         | settofalse). |
| --------------------------- | ------- | ------------ |
| Modifying                   | Systems | of Units     |
MethodIntroduced:
• IpfcUnitSystem.Delete()
ThemethodIpfcUnitSystem.Delete() deletesacustom-definedsystem
ofunits.
Note
Youcandelete onlya custom-definedsystemof unitsandnotastandard
| systemof | units.       |          |
| -------- | ------------ | -------- |
| Creating | a New System | of Units |
MethodIntroduced:
• IpfcSolid.CreateUnitSystem()
ThemethodIpfcSolid.CreateUnitSystem() createsa newsystemof
unitsin themodelbasedon thespecifiedname,thetypeofunitsystem givenby
theIpfcUnitSystemTypeobject,andthe typesof unitsspecifiedby the
IpfcUnitssequence tousefor eachofthebase measurementtypes(length,
| forceor mass,andtemperature). |          |             |
| ----------------------------- | -------- | ----------- |
| Conversion                    | to a New | Unit System |
MethodsandPropertiesIntroduced:
• IpfcSolid.SetPrincipalUnits()
• CCpfcUnitConversionOptions.Create()
• IpfcUnitConversionOptions.DimensionOption
Solid 163

• IpfcUnitConversionOptions.IgnoreParamUnits
ThemethodIpfcSolid.SetPrincipalUnits()
changestheprincipal
systemofunitsassignedto thesolidmodelbasedonthe unitconversionoptions
| specifiedby                        | theIpfcUnitConversionOptions |     | object.The | method |
| ---------------------------------- | ---------------------------- | --- | ---------- | ------ |
| IpfcUnitConversionOptions.Create() |                              |     | createsthe |        |
IpfcUnitConversionOptionsobjectcontainingtheunitconversion
| optionslisted                   | below. |               |     |     |
| ------------------------------- | ------ | ------------- | --- | --- |
| Thetypesofunitconversionoptions |        | areasfollows: |     |     |
• DimensionOption—Use theoption whileconvertingthe dimensionsof
themodel.
| Usethe                                | propertyIpfcUnitConversionOptions.DimensionOption |                 |                    |     |
| ------------------------------------- | ------------------------------------------------- | --------------- | ------------------ | --- |
| tomodifythis                          | option.                                           |                 |                    |     |
| Thisoption                            | canbe ofthe                                       | followingtypes: |                    |     |
| ○ EpfcUNITCONVERT_SAME_DIMS—Specifies |                                                   |                 | thatunitconversion |     |
occursby interpretingtheunitvalue inthe newunitsystem. Forexample,
| 1inchwillequalto                      |     | 1millimeter. |                        |     |
| ------------------------------------- | --- | ------------ | ---------------------- | --- |
| ○ EpfcUNITCONVERT_SAME_SIZE—Specifies |     |              | thatunitconversionwill |     |
occurby convertingtheunitvalue inthe newunitsystem. Forexample,1
| inchwillequalto |     | 25.4millimeters. |     |     |
| --------------- | --- | ---------------- | --- | --- |
• IgnoreParamUnits—Thisbooleanattribute determineswhetheror not
ignorethe parameterunits. Ifitisnullortrue, parametervalues andunitsdo
notchangewhenthe unitsystem ischanged.Ifitis false,parameter unitsare
| convertedaccordingto                       |          | therule. |     |               |
| ------------------------------------------ | -------- | -------- | --- | ------------- |
| Usethe                                     | property |          |     |               |
| IpfcUnitConversionOptions.IgnoreParamUnits |          |          |     | to modifythis |
attribute.
| Mass | Properties |     |     |     |
| ---- | ---------- | --- | --- | --- |
MethodIntroduced:
• IpfcSolid.GetMassProperty()
• IpfcSolid.GetMassPropertyWithDensity()
• IpfcAssembly.GetMassPropertyByCompPath()
Thefunction IpfcSolid.GetMassProperty()providesinformationabout
thedistributionofmass in thepartor assembly.Itcan providetheinformation
relativeto acoordinatesystem datum,whichyouname,or thedefaultoneif you
providenullas thename..Itreturns anobjectcontainingthe followingfields:
• Thevolume.
| 164 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

| • Thesurface | area. |     |     |
| ------------ | ----- | --- | --- |
• Thedensity. Thedensityvalue is1.0,unlessamaterial hasbeenassigned.
• Themass.
| • Thecenterof | gravity(COG). |     |     |
| ------------- | ------------- | --- | --- |
| • Theinertia  | matrix.       |     |     |
| • Theinertia  | tensor.       |     |     |
| • Theinertia  | abouttheCOG.  |     |     |
• Theprincipal momentsof inertia(the eigenvalues oftheCOGinertia).
| • Theprincipal | axes(the eigenvectorsofthe | COGinertia). |     |
| -------------- | -------------------------- | ------------ | --- |
ThepropertyIpfcSolid.GetMassPropertyWithDensity() calculates
themasspropertiesof apartor an assemblyin thespecifiedcoordinatesystem.
Theinputargumentsfollow:
• CoordSysName—Nameof thecoordinatesystem.Ifthis isNull, themethod
usesthedefaultcoordinatesystem.
• DensityOpt—Valueof thedensityflag specifiedusing theenumerated data
| typeIpfcMPDensityUseandthevalid |     | valuesareas | follows: |
| ------------------------------- | --- | ----------- | -------- |
○ EpfcMP_DENSITY_DEFAULT—Calculate themass propertiesusingthe
materialdensity.
EpfcMP_DENSITY_USE_ALWAYS—Calculatethe
| ○                                         |     |                             | massproperties |
| ----------------------------------------- | --- | --------------------------- | -------------- |
| using thespecifieddensity,evenif          |     | materialhasadefineddensity. |                |
| ○ EpfcMP_DENSITY_USE_IF_MISSING—Calculate |     |                             | massproperties |
using specifieddensity, evenifmaterial doesnothaveadefineddensity.
• density—Densityusedwhilecalculatingmasspropertiesdependingon the
| valuespecifiedfor | theinputargumentDensityOpt. |     |     |
| ----------------- | --------------------------- | --- | --- |
ThemethodIpfcAssembly.GetMassPropertyByCompPath()
calculatesthe masspropertiesof asolidthatisreferencedby thespecified
coordinatesystem selection,usingthe respectivecomponentpaths.The input
argumentsfollow:
• CompPath—Componentpathof thesolid. Ifthisis null,thetop assemblyis
referred.
• CsysItem—Coordinatesystem modelitem.Ifthis isnull,default coordinate
| systemis | referred. |     |     |
| -------- | --------- | --- | --- |
• CsysPath—Componentpathofthe coordinatesystem.If thisisnull, default
| coordinatesystem | isreferred. |     |     |
| ---------------- | ----------- | --- | --- |
Solid 165

| Example                                  | Code: Retrieving           | a Mass | Property        | Object |
| ---------------------------------------- | -------------------------- | ------ | --------------- | ------ |
| Thesamplecodein                          | thefilepfcSolidExamples.vb |        | locatedat<creo_ |        |
| vbapi_loadpoint>/vbapi_examplesretrieves |                            |        | aMassProperty   |        |
objectfrom aspecifiedsolidmodel. Thesolid'smass,volume,andcenterof
gravitypointarethenprinted.
Annotations
MethodsandPropertiesIntroduced:
• IpfcNote.Lines
• IpfcNote.GetText()
• IpfcNote.URL
• IpfcNote.Display()
• IpfcNote.Delete()
• IpfcNote.GetOwner()
3Dmodelnotesareinstanceof ModelItemobjects.Theycan belocated and
accessedusingmethods thatlocatemodelitemsin solidmodels,and downcastto
| theNoteinterface | tousethe methodsin | thissection. |     |     |
| ---------------- | ------------------ | ------------ | --- | --- |
ThepropertyIpfcNote.Lines returnsthetextcontained inthe 3Dmodel
note.
ThemethodIpfcNote.GetText() returnsthetextof thesolidmodelnote.If
yousettheparameterGiveParametersAsNames to TRUE,thenthe textdisplays
theparametercallouts with ampersands(&).Ifyousettheparameterto FALSE,
thenthe textdisplaysthe parametervalues withno callout information.
ThepropertyIpfcNote.URLreturns theURL storedinthe 3Dmodelnote.
| ThemethodIpfcNote.Display() |     | forcesthedisplayofthe |     | modelnote. |
| --------------------------- | --- | --------------------- | --- | ---------- |
| ThemethodIpfcNote.Delete()  |     | deletesa modelnote.   |     |            |
ThemethodIpfcNote.GetOwner() returns thesolidmodelownerof the
note.
| Cross | Sections |     |     |     |
| ----- | -------- | --- | --- | --- |
MethodsIntroduced:
• IpfcSolid.ListCrossSections()
• IpfcSolid.GetCrossSection()
• IpfcXSection.GetName()
• IpfcXSection.SetName()
| 166 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

• IpfcXSection.GetXSecType()
• IpfcXSection.Delete()
• IpfcXSection.Display()
• IpfcXSection.Regenerate()
ThemethodIpfcSolid.ListCrossSections() returnsasequence of
crosssectionobjectsrepresentedby theXsectioninterface.The method
IpfcSolid.GetCrossSection()searchesfora crosssectiongivenits
name.
ThemethodIpfcXSection.GetName() returnsthe nameofthe crosssection
inCreo Parametric.ThemethodIpfcXSection.SetName() modifies the
crosssectionname.
ThemethodIpfcXSection.GetXSecType() returnsthe typeofcross
section,thatis planaror offset,andthetypeof itemintersectedby thecross
section.
ThemethodIpfcXSection.Delete() deletes acrosssection.
ThemethodIpfcXSection.Display() forcesadisplayofthe crosssection
inthe window.
ThemethodIpfcXSection.Regenerate() regeneratesa crosssection.
Materials
theVBAPI enablesyouto programmaticallyaccessthematerial typesand
propertiesof parts.Using themethodsand propertiesdescribed in thefollowing
sections,youcan performthefollowingactions:
• Createordeletematerials
• Setthecurrentmaterial
• Accessandmodifythematerial typesand properties
MethodsandPropertiesIntroduced:
• IpfcMaterial.Save()
• IpfcMaterial.Delete()
• IpfcPart.CurrentMaterial
• IpfcPart.ListMaterials()
• IpfcPart.CreateMaterial()
• IpfcPart.RetrieveMaterial()
ThemethodIpfcMaterial.Save() writesto amaterial filethatcanbe
importedintoany CreoParametricpart.
Solid 167

| ThemethodIpfcMaterial.Delete() |     |     |     | removes | materialfrom | thepart. |
| ------------------------------ | --- | --- | --- | ------- | ------------ | -------- |
ThepropertyIpfcPart.CurrentMaterial
|            |           |     |     |     | returnsandsetsthe | material |
| ---------- | --------- | --- | --- | --- | ----------------- | -------- |
| assignedto | the part. |     |     |     |                   |          |
Note
Bydefault,whileassigninga materialto asheetmetal part,theproperty
| IpfcPart.CurrentMaterial |     |     |     | modifiesthe | valuesof thesheetmetal |     |
| ------------------------ | --- | --- | --- | ----------- | ---------------------- | --- |
propertiessuchas Yfactorandbend tableaccordingto thematerial file
definition.This modificationtriggersaregeneration andamodification ofthe
developedlengthcalculationsofthe sheetmetalpart.However,youcanavoid
thisbehavior bysetting thevalueof theconfigurationoptionmaterial_
| update_smt_bend_table |     |     |     | to never_replace |     |     |
| --------------------- | --- | --- | --- | ---------------- | --- | --- |
The propertyIpfcPart.CurrentMaterialmaychangethemodel
| display,if          | thenewmaterial |                         | hasadefaultappearanceassignedto |              |              | it.  |
| ------------------- | -------------- | ----------------------- | ------------------------------- | ------------ | ------------ | ---- |
| The propertymayalso |                | changethefamilytable,if |                                 |              | theparameter | PTC_ |
| MATERIAL_NAME       |                | isapart                 | ofthe                           | familytable. |              |      |
ThemethodIpfcPart.ListMaterials() returnsa listof thematerials
| availablein | thepart. |     |     |     |     |     |
| ----------- | -------- | --- | --- | --- | --- | --- |
ThemethodIpfcPart.CreateMaterial() createsanewemptymaterialin
thespecifiedpart.
ThemethodIpfcPart.RetrieveMaterial() importsamaterial fileinto
| thepart.The                          | nameofthe | fileread | canbe | as either:  |     |     |
| ------------------------------------ | --------- | -------- | ----- | ----------- | --- | --- |
| • <name>.mtl—Specifiesthenewmaterial |           |          |       | fileformat. |     |     |
• <name>.mat—Specifiesthematerial fileformatprior to Pro/ENGINEER
Wildfire3.0.
| Ifthematerial | isnotalreadyin |     | thepart | database, |     |     |
| ------------- | -------------- | --- | ------- | --------- | --- | --- |
IpfcPart.RetrieveMaterial()addsthe materialto thedatabase after
readingthematerialfile. Ifthematerial isalreadyin the database,thefunction
replacesthematerial propertiesinthe databasewith thosecontainedin the
materialfile.
| Accessing | Material |     | Types |     |     |     |
| --------- | -------- | --- | ----- | --- | --- | --- |
PropertiesIntroduced:
• IpfcMaterial.StructuralMaterialType
• IpfcMaterial.ThermalMaterialType
• IpfcMaterial.SubType
| 168 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

• IpfcMaterial.PermittedSubTypes
• IpfcMaterial.FluidMaterialType
| ThepropertyIpfcMaterial.StructuralMaterialType |     |     |     | setsthe |     |
| ---------------------------------------------- | --- | --- | --- | ------- | --- |
materialtypeforthe structuralpropertiesofthematerial. Thematerial typesareas
follows:
EpfcMTL_ISOTROPIC—Specifies
| •   |     | aamaterial | withan | infinitenumberof |     |
| --- | --- | ---------- | ------ | ---------------- | --- |
planesof materialsymmetry, makingthepropertiesequal inall directions.
• EpfcMTL_ORTHOTROPIC—Specifies amaterial withsymmetryrelativeto
threemutuallyperpendicularplanes.
| • EpfcMTL_TRANSVERSELY_ISOTROPIC—Specifies |     |     |     | amaterial | with |
| ------------------------------------------ | --- | --- | --- | --------- | ---- |
rotationalsymmetryaboutan axis.Thepropertiesareequal forall directions
inthe planeof isotropy.
| • EpfcMTL_FLUID—Specifiesa                  | materialwith |     | fluidproperties. |         |          |
| ------------------------------------------- | ------------ | --- | ---------------- | ------- | -------- |
| ThepropertyIpfcMaterial.ThermalMaterialType |              |     |                  | setsthe | material |
typefor thethermalpropertiesofthe material.Thematerial typesareasfollows:
EpfcMTL_ISOTROPIC—Specifies
| •   |     | amaterialwith |     | an infinitenumberof |     |
| --- | --- | ------------- | --- | ------------------- | --- |
planesof materialsymmetry, makingthepropertiesequal inall directions.
• EpfcMTL_ORTHOTROPIC—Specifies amaterial withsymmetryrelativeto
threemutuallyperpendicularplanes.
| • EpfcMTL_TRANSVERSELY_ISOTROPIC—Specifies |     |     |     | amaterial | with |
| ------------------------------------------ | --- | --- | --- | --------- | ---- |
rotationalsymmetryaboutan axis.Thepropertiesareequal forall directions
inthe planeof isotropy.
| • EpfcMTL_FLUID—Specifiesa | materialwith |     | fluidproperties. |     |     |
| -------------------------- | ------------ | --- | ---------------- | --- | --- |
ThepropertyIpfcMaterial.SubType setsthesubtype forthe EpfcMTL_
| ISOTROPICmaterial type. |     |     |     |     |     |
| ----------------------- | --- | --- | --- | --- | --- |
Usethe propertyIpfcMaterial.PermittedSubTypes toretrieve alist of
| thepermittedstringvaluesfor | thematerial | subtype. |     |     |     |
| --------------------------- | ----------- | -------- | --- | --- | --- |
ThepropertyIpfcMaterial.FluidMaterialType setsthe materialtype
forthefluidpropertiesof thematerial.The materialtypesareas follows:
• EpfcMTL_ISOTROPIC—Specifies amaterialwith an infinitenumberof
planesof materialsymmetry, makingthepropertiesequal inall directions.
• EpfcMTL_ORTHOTROPIC—Specifies amaterial withsymmetryrelativeto
threemutuallyperpendicularplanes.
EpfcMTL_TRANSVERSELY_ISOTROPIC—Specifies
| •   |     |     |     | amaterial | with |
| --- | --- | --- | --- | --------- | ---- |
rotationalsymmetryaboutan axis.Thepropertiesareequal forall directions
inthe planeof isotropy.
| • EpfcMTL_FLUID—Specifiesa | materialwith |     | fluidproperties. |     |     |
| -------------------------- | ------------ | --- | ---------------- | --- | --- |
Solid 169

| Accessing | Material | Properties |     |     |
| --------- | -------- | ---------- | --- | --- |
Themethodsand propertieslisted in thissectionenableyouto accessmaterial
properties.
MethodsandPropertiesIntroduced:
• CCpfcMaterialProperty.Create()
• IpfcMaterial.GetPropertyValue()
• IpfcMaterial.SetPropertyValue()
• IpfcMaterial.SetPropertyUnits()
• IpfcMaterial.RemoveProperty()
• IpfcMaterial.Description
• IpfcMaterial.FatigueType
• IpfcMaterial.PermittedFatigueTypes
• IpfcMaterial.FatigueMaterialType
• IpfcMaterial.PermittedFatigueMaterialTypes
• IpfcMaterial.FatigueMaterialFinish
• IpfcMaterial.PermittedFatigueMaterialFinishes
• IpfcMaterial.FailureCriterion
• IpfcMaterial.PermittedFailureCriteria
• IpfcMaterial.Hardness
• IpfcMaterial.HardnessType
• IpfcMaterial.Condition
• IpfcMaterial.BendTable
• IpfcMaterial.CrossHatchFile
• IpfcMaterial.MaterialModel
• IpfcMaterial.PermittedMaterialModels
• IpfcMaterial.ModelDefByTests
| ThemethodCCpfcMaterialProperty.Create() |     |     | createsa | newinstance |
| --------------------------------------- | --- | --- | -------- | ----------- |
| ofamaterial propertyobject.             |     |     |          |             |
Allnumericalmaterial propertiesareaccessedusing thesamesetof APIs.You
mustprovidea propertytypeto indicatethe propertyyouwantto read ormodify.
| ThemethodIpfcMaterial.GetPropertyValue() |           |     | returnsthe | valueand |
| ---------------------------------------- | --------- | --- | ---------- | -------- |
| theunitsofthematerial                    | property. |     |            |          |
methodIpfcMaterial.SetPropertyValue()to
| Usethe |     |     | setthevalue | and |
| ------ | --- | --- | ----------- | --- |
unitsof thematerial property.Ifthe propertytypedoesnotexistfor thematerial,
| thenthis methodcreatesit. |     |                                 |     |     |
| ------------------------- | --- | ------------------------------- | --- | --- |
| 170                       |     | Creo®ParametricVBAPIUser’sGuide |     |     |

| Usethe methodIpfcMaterial.SetPropertyUnits()to |     |     | settheunitsof |     |
| ---------------------------------------------- | --- | --- | ------------- | --- |
thematerialproperty.
| Usethe methodIpfcMaterial.RemoveProperty()to |     | removethe |     |     |
| -------------------------------------------- | --- | --------- | --- | --- |
materialproperty.
Materialpropertiesthatarenon-numeric canbe accessedusing thefollowing
properties.
ThepropertyIpfcMaterial.Description setsthedescriptionstringfor
thematerial.
ThepropertyIpfcMaterial.FatigueType
|     |     | andsetsthe validfatiguetype |     |     |
| --- | --- | --------------------------- | --- | --- |
forthematerial.
Usethe propertyIpfcMaterial.PermittedFatigueTypes to getalist of
thepermittedstringvaluesfor thefatiguetype.
| ThepropertyIpfcMaterial.FatigueMaterialType |             | setsthe  | classof |     |
| ------------------------------------------- | ----------- | -------- | ------- | --- |
| materialwhendeterminingthe                  | effectofthe | fatigue. |         |     |
Usethe propertyIpfcMaterial.PermittedFatigueMaterialTypes
toretrieve alist ofthepermittedstringvalues forthe fatiguematerialtype.
| ThepropertyIpfcMaterial.FatigueMaterialFinish |     |     | setsthe typeof |     |
| --------------------------------------------- | --- | --- | -------------- | --- |
| surfacefinishfor thefatiguematerial.          |     |     |                |     |
Usethe property
| IpfcMaterial.PermittedFatigueMaterialFinishes |                               |     | toretrieve | a   |
| --------------------------------------------- | ----------------------------- | --- | ---------- | --- |
| listofpermittedstringvalues                   | forthe fatiguematerialfinish. |     |            |     |
ThepropertyIpfcMaterial.FailureCriterion setsthereductionfactor
forthefailurestrength ofthematerial.This factoris usedto reducethe endurance
limitofthematerial to accountfor unmodeledstress concentrations,suchasthose
foundin welds.
Usethe propertyIpfcMaterial.PermittedFailureCriteriato
retrievealist ofpermittedstringvalues forthe materialfailurecriterion.
ThepropertyIpfcMaterial.Hardness setsthe hardnessfor thespecified
material.
ThepropertyIpfcMaterial.HardnessType
|     |     | setsthehardnesstypefor |     | the |
| --- | --- | ---------------------- | --- | --- |
specifiedmaterial.
ThepropertyIpfcMaterial.Condition setsthecondition forthespecified
material.
ThepropertyIpfcMaterial.BendTable setsthebend tablefor thespecified
material.
ThepropertyIpfcMaterial.CrossHatchFile setsthefilecontainingthe
| crosshatchpatternfor | thespecifiedmaterial | .   |     |     |
| -------------------- | -------------------- | --- | --- | --- |
| Solid                |                      |     |     | 171 |

ThepropertyIpfcMaterial.MaterialModel setsthetypeof hyperelastic
| isotropicmaterial | model. |     |     |
| ----------------- | ------ | --- | --- |
Usethe propertyIpfcMaterial.PermittedMaterialModels to retrieve
| alistof thepermittedstringvaluesfor     |     | thematerial | model.               |
| --------------------------------------- | --- | ----------- | -------------------- |
| ThepropertyIpfcMaterial.ModelDefByTests |     |             | determineswhetherthe |
hyperelasticisotropic materialmodelhasbeendefinedusingexperimental datafor
stressandstrain.
| Accessing | User-defined | Material | Properties |
| --------- | ------------ | -------- | ---------- |
Materialspermitassignmentof user-definedparameters.These parametersallow
| youtoplace | non-standardpropertieson | agiven | material.Therefore |
| ---------- | ------------------------ | ------ | ------------------ |
IpfcMaterialisachild ofIpfcParameterOwner,whichprovidesaccess
touser-defined parametersandpropertiesof materialsthrough themethods in that
interface.
| 172 |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | ------------------------------- |

12
Solid Bodies
SolidBodyInformation.............................................................................................174
Thischapterdescribes theVB APIfunctions thataccessthefunctions ofaCreo
Parametricpartwith multiplesolidbodies.
173

Solid Body Information
Thestateof thebodyis derivedfromthe featuresandgeometryin whichthebody
iscreated.
Methodsintroduced:
• IpfcSolidBody.GetBodyState()
• IpfcSolidBody.IsConstruction()
• IpfcSolidBody.ListSurfaces()
• IpfcSolid.GetDefaultBody()
• IpfcSolidBody.GetFeatures()
• IpfcSolidBody.GetOutline()
• IpfcSolidBody.IsSheetmetal()
• IpfcSolidBody.GetDensity()
• IpfcSolidBody.GetMassProperty()
ThemethodIpfcSolidBody.GetBodyState() specifies thestateof the
bodyusing theenumeratedtypeEpfcSolidBodyState.The validvalues are:
• EpfcBODY_STATE_MISSING
• EpfcBODY_STATE_CONSUMED
• EpfcBODY_STATE_NO_CONTR_FEAT
• EpfcBODY_STATE_NO_GEOMETRY
• EpfcBODY_STATE_ACTIVE
ThemethodIpfcSolidBody.IsConstruction() checks ifthe bodyis a
constructionbody. Themethodreturns aTrueif solidbody isa construction
body.
ThemethodIpfcSolidBody.ListSurfaces() lists allthe surfacesin the
specifiedbody
ThemethodIpfcSolid.GetDefaultBody() returnsthe defaultbodyin the
specifiedsolid.
ThemethodIpfcSolidBody.GetFeatures() liststhe featuresthatare
associatedwiththe specifiedbody.
Usethe methodIpfcSolidBody.GetOutline() toretrieve the
regenerationoutlineof asolidbody, with respectto thebase coordinatesystem
orientation.This outlinedefines theboundary boxofthebody.
ThemethodIpfcSolidBody.IsSheetmetal() checks ifthe specified
bodyisan activesheetmetal body.
ThemethodIpfcSolidBody.GetDensity() determinesthedensityof the
body.
174 Creo®ParametricVBAPIUser’sGuide

Usethe methodIpfcSolidBody.GetMassProperty()to retrieve the
masspropertiesofa bodyin thespecifiedcoordinatesystem.
SolidBodies 175

13
Windows and Views
Windows.................................................................................................................177
EmbeddedBrowser.................................................................................................179
Views.....................................................................................................................180
CoordinateSystemsandTransformations.................................................................181
TheVB APIprovidesaccessto Creo Parametricwindowsandsavedviews. This
| chapterdescribesthemethods | thatprovidethis | access.                         |
| -------------------------- | --------------- | ------------------------------- |
| 176                        |                 | Creo®ParametricVBAPIUser’sGuide |

Windows
Thissectiondescribesthe VBAPI methodsthataccessWindowobjects.The
topicsareasfollows:
• GettingaWindowObjectonpage177
| • WindowOperationson |          | page178 |     |     |     |
| -------------------- | -------- | ------- | --- | --- | --- |
| Getting              | a Window | Object  |     |     |     |
MethodsandPropertyIntroduced:
• IpfcBaseSession.CurrentWindow
• IpfcBaseSession.CreateModelWindow()
• IpfcModel.Display()
• IpfcBaseSession.ListWindows()
• IpfcBaseSession.GetWindow()
• IpfcBaseSession.OpenFile()
• IpfcBaseSession.GetModelWindow()
ThepropertyIpfcBaseSession.CurrentWindow providesaccessto the
| currentactivewindow                          | inCreo | Parametric. |            |             |     |
| -------------------------------------------- | ------ | ----------- | ---------- | ----------- | --- |
| ThemethodIpfcBaseSession.CreateModelWindow() |        |             |            | createsanew |     |
| windowthatcontainsthemodelthatwas            |        |             | passedasan | argument.   |     |
Note
Youmustcallthemethod IpfcModel.Display() forthe modelgeometry
| to bedisplayedin | thewindow. |     |     |     |     |
| ---------------- | ---------- | --- | --- | --- | --- |
Usethe methodIpfcBaseSession.ListWindows()to geta listof allthe
| currentwindows                       | insession. |     |     |                 |         |
| ------------------------------------ | ---------- | --- | --- | --------------- | ------- |
| ThemethodIpfcBaseSession.GetWindow() |            |     |     | getsthehandleto | awindow |
givenitsintegeridentifier.
ThemethodIpfcBaseSession.OpenFile() returnsthe handletoa newly
| createdwindowthatcontainsthe |     | openedmodel. |     |     |     |
| ---------------------------- | --- | ------------ | --- | --- | --- |
| WindowsandViews              |     |              |     |     | 177 |

Note
Ifamodelis alreadyopenin a windowthemethodreturns ahandletothe
window.
ThemethodIpfcBaseSession.GetModelWindow() returns thehandleto
| thewindowthatcontainsthe | openedmodel, | ifitis | displayed. |     |
| ------------------------ | ------------ | ------ | ---------- | --- |
Window Operations
MethodsandPropertiesIntroduced:
• IpfcWindow.Height
• IpfcWindow.Width
• IpfcWindow.XPos
• IpfcWindow.YPos
• IpfcWindow.GraphicsAreaHeight
• IpfcWindow.GraphicsAreaWidth
• IpfcWindow.Clear()
• IpfcWindow.Repaint()
• IpfcWindow.Refresh()
• IpfcWindow.Close()
• IpfcWindow.Activate()
• IpfcWindow.GetId
• IpfcBaseSession.FlushCurrentWindow()
| ThepropertiesIpfcWindow.Height, |     | IpfcWindow.Width, |     |     |
| ------------------------------- | --- | ----------------- | --- | --- |
IpfcWindow.XPos,andIpfcWindow.YPos retrieve theheight, width,x-
position,andy-positionofthewindow respectively. Thevalues ofthese
| parametersarenormalized                    | from0to | 1.  |                     |     |
| ------------------------------------------ | ------- | --- | ------------------- | --- |
| ThepropertiesIpfcWindow.GraphicsAreaHeight |         |     | and                 |     |
| IpfcWindow.GraphicsAreaWidthretrieve       |         |     | theheightandwidthof | the |
CreoParametricgraphicsareawindowwithoutthe borderrespectively. Thevalues
oftheseparametersarenormalizedfrom 0to1.For both thewindowandgraphics
areasizes, iftheobjectoccupies thewholescreen,the windowsizereturned is1.
Forexample,ifthe screenis1024pixelswide andthegraphicsareais 512
| pixels,thenthewidthof       | thegraphicsareawindowis |                     | returnedas0.5.                  |            |
| --------------------------- | ----------------------- | ------------------- | ------------------------------- | ---------- |
| ThemethodIpfcWindow.Clear() |                         | removesgeometryfrom |                                 | thewindow. |
| 178                         |                         |                     | Creo®ParametricVBAPIUser’sGuide |            |

| BothIpfcWindow.Repaint() | andIpfcWindow.Refresh() | repaint |
| ------------------------ | ----------------------- | ------- |
solidgeometry.However, theRefresh() methoddoesnotremovehighlights
fromthescreen andisused primarilyto removetemporarygeometryentitiesfrom
thescreen.
Usethe methodIpfcWindow.Close()to closethewindow.Ifthe current
windowistheoriginal windowcreated whenCreo Parametricstarted,this method
| clearsthewindow.Otherwise,itremovesthewindow | fromthe | screen. |
| -------------------------------------------- | ------- | ------- |
ThemethodIpfcWindow.Activate()
activates awindow. This functionis
availableonly in theasynchronousmode.
| ThemethodpfcWindow.Window.GetId | retrieves theIDof | theCreo |
| ------------------------------- | ----------------- | ------- |
Parametricwindow.
| ThemethodIpfcBaseSession.FlushCurrentWindow() |                | flushesthe |
| --------------------------------------------- | -------------- | ---------- |
| pendingdisplaycommands onthe                  | currentwindow. |            |
Note
It isrecommendedto callthismethod onlyaftercompleting allthedisplay
operations.Excessiveuseofthismethod willcausemajor slowdownof
systems runningon WindowsVistaandWindows7.
Embedded Browser
MethodsIntroduced:
• IpfcWindow.GetURL()
• IpfcWindow.SetURL()
• IpfcWindow.GetBrowserSize()
• IpfcWindow.SetBrowserSize()
| ThemethodsIpfcWindow.GetURL() | and IpfcWindow.SetURL() |     |
| ----------------------------- | ----------------------- | --- |
enablesyoutofind andchangetheURL displayedin the embeddedbrowserin the
CreoParametricwindow.
WindowsandViews 179

ThemethodsIpfcWindow.GetBrowserSize() and
IpfcWindow.SetBrowserSize()enablesyouto findandchangethe size
| oftheembedded | browserin | the | CreoParametricwindow. |     |     |
| ------------- | --------- | --- | --------------------- | --- | --- |
Note
The methodsIpfcWindow.GetBrowserSize()and
| IpfcWindow.SetBrowserSize()are |                     |     |     | notsupportedif | thebrowseris |
| ------------------------------ | ------------------- | --- | --- | -------------- | ------------ |
| open                           | ina separatewindow. |     |     |                |              |
Views
Thissectiondescribesthe VBAPI methodsthataccessIpfcView objects.The
topicsareasfollows:
• GettingaViewObjectonpage180
| • ViewOperationson |        | page181 |     |     |     |
| ------------------ | ------ | ------- | --- | --- | --- |
| Getting            | a View | Object  |     |     |     |
MethodsIntroduced:
• IpfcViewOwner.RetrieveView()
• IpfcViewOwner.GetView()
• IpfcViewOwner.ListViews()
• IpfcViewOwner.GetCurrentView()
theinterfaceIpfcViewOwner.
| Anysolidmodelinheritsfrom |     |                 |     |     | This willenable |
| ------------------------- | --- | --------------- | --- | --- | --------------- |
| youtousethesemethodson    |     | anysolidobject. |     |     |                 |
ThemethodIpfcViewOwner.RetrieveView() setsthe currentview to the
| orientationpreviouslysavedwitha |                               |                | specifiedname. |                 |        |
| ------------------------------- | ----------------------------- | -------------- | -------------- | --------------- | ------ |
| Usethe                          | methodIpfcViewOwner.GetView() |                |                | to getahandleto | anamed |
| viewwithoutmakingany            |                               | modifications. |                |                 |        |
ThemethodIpfcViewOwner.ListViews() returnsalist ofall theviews
| previouslysavedin              |                   | themodel. |                   |                    |           |
| ------------------------------ | ----------------- | --------- | ----------------- | ------------------ | --------- |
| FromCreo                       | Parametric2.0M120 |           | onward,themethod, |                    |           |
| IpfcViewOwner.GetCurrentView() |                   |           |                   | hasbeendeprecated. | Themethod |
returnsaview handlethatrepresents thecurrentorientation.Although thisview
doesnothaveaname,youcan usethis viewto find ormodifythecurrent
orientation.
| 180 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

View Operations
MethodsandPropertiesIntroduced:
• IpfcView.Name
• IpfcView.IsCurrent
• IpfcView.Reset()
• IpfcViewOwner.SaveView()
| Toget thenameof | aview givenitsidentifier, | usethe property |
| --------------- | ------------------------- | --------------- |
IpfcView.Name.
ThepropertyIpfcView.IsCurrent determinesiftheViewobjectrepresents
thecurrentview.
ThemethodIpfcView.Reset() restoresthe currentview to thedefaultview.
| Tostore thecurrentview | underthespecifiedname,callthemethod |     |
| ---------------------- | ----------------------------------- | --- |
IpfcViewOwner.SaveView().
| Coordinate | Systems | and Transformations |
| ---------- | ------- | ------------------- |
hissectiondescribes thevarious coordinatesystems usedby CreoParametricand
accessiblefrom theVB APIandhowto transformfrom onecoordinatesystemto
another.
| Coordinate           | Systems                                 |     |
| -------------------- | --------------------------------------- | --- |
| CreoParametricandthe | VBAPI usethefollowingcoordinatesystems: |     |
• SolidCoordinateSystemon page182
• ScreenCoordinateSystem onpage182
• WindowCoordinateSystem onpage182
• DrawingCoordinateSystem onpage183
| • DrawingViewCoordinateSystemon |     | page183 |
| ------------------------------- | --- | ------- |
| • AssemblyCoordinateSystemon    |     | page183 |
• DatumCoordinateSystem onpage183
| • SectionCoordinateSystemon | page183 |     |
| --------------------------- | ------- | --- |
Thefollowingsectionsdescribeeachofthesecoordinatesystems.
WindowsandViews 181

Solid Coordinate System
Thesolidcoordinatesystemisthe three-dimensional,Cartesiancoordinatesystem
usedto describethegeometryof aCreoParametric solidmodel. Ina part,the
solidcoordinatesystem describesthegeometryof thesurfacesand edges.In an
assembly, thesolidcoordinatesystem alsodescribes thelocationsandorientations
oftheassembly members.
You canvisualizethe solidcoordinatesystem inCreo Parametricby creatinga
coordinatesystem datumwith theoption Default.Distances measuredin solid
coordinatescorrespond to thevaluesof dimensionsas seenby theCreo Parametric
user.
Solidcoordinatesareused bythe VBAPI forall themethods thatlook at
geometryandmostofthemethods thatdrawthree-dimensionalgraphics.
Screen Coordinate System
Thescreencoordinatesystemis two-dimensionalcoordinatesystemthatdescribes
locationsina CreoParametricwindow.This isan intermediatecoordinatesystem
afterwhich thescreenpointsaretransformed to screenpixels.Allthe modelsare
firstmappedto thescreencoordinatesystem.When theuserzoomsor pansthe
view, thescreencoordinatesystemfollows thedisplayofthe solid,soa particular
pointonthe solidalways mapstothe samescreencoordinate.The mapping
changesonly whenthe vieworientation ischanged.
Screencoordinatesareusedby someof thegraphicsmethods,the mouseinput
methods,andallmethods thatdraw graphicsormanipulateitemson adrawing.
Window Coordinate System
Thewindowcoordinatesystemis similarto thescreencoordinatesystem.After
mappingthe modelsto thescreen coordinatesystem, theyare mappedtothe
windowcoordinatebefore beingdrawn toscreenpixels basedonscreen
resolution.When panorzoomvalues areappliedto thecoordinatesin thescreen
coordinatesystem,theyresultin windowcoordinates. Whenan objectisfirst
displayedin awindow,or theoption View▶▶Refitis used,thescreenandwindow
coordinatesarethe same.
Windowcoordinatesareneededonlyif youneedtotakeaccount ofzoomandpan
—forexample,tofind outwhetherapointon thesolidis visibleinthe window,or
todraw two-dimensionaltextina particularwindowlocation,regardless ofpan
andzoom.
182 Creo®ParametricVBAPIUser’sGuide

| Drawing | Coordinate | System |     |
| ------- | ---------- | ------ | --- |
Thedrawing coordinatesystem isa two-dimensionalsystemthatdescribesthe
locationon adrawing relativeto thebottom,leftcorner,andmeasured in drawing
units.For example,on aU.S. letter-sized,landscape-formatdrawing sheetthat
| usesinches, | thetop,right-corneris(11,8.5)in |     | drawing coordinates. |
| ----------- | ------------------------------- | --- | -------------------- |
TheVB APImethods andpropertiesthatmanipulatedrawings generallyuse
screencoordinates.
| Drawing | View Coordinate | System |     |
| ------- | --------------- | ------ | --- |
Thedrawing viewcoordinatesystem isusedto describethe locationsofentitiesin
| adrawing | view.      |        |     |
| -------- | ---------- | ------ | --- |
| Assembly | Coordinate | System |     |
Anassemblyhasitsowncoordinatesystemthatdescribes thepositionsand
orientationsofthe memberparts, subassemblies,andthegeometryofdatum
| featurescreatedin | theassembly. |     |     |
| ----------------- | ------------ | --- | --- |
Whenanassembly isretrievedintomemory eachmemberis alsoloadedand
continuesto useitsownsolidcoordinatesystem to describeitsgeometry.
Thisisimportantwhen youareanalyzingthe geometryofasubassembly and
wantto extractor displaythe resultsrelativeto thecoordinatesystemof theparent
assembly.
| Datum | Coordinate | System |     |
| ----- | ---------- | ------ | --- |
Acoordinatesystem datumcan be createdanywherein anypart orassembly,and
representsauser-defined coordinatesystem. Itis oftenarequirementin aVB API
| applicationto | describegeometryrelativeto |        | suchadatum. |
| ------------- | -------------------------- | ------ | ----------- |
| Section       | Coordinate                 | System |             |
Everysketchhasacoordinatesystemused tolocate entitiesin thatsketch.
Sketchesusedin featureswillusea coordinatesystem differentfrom thatofthe
solidmodel.
Transformations
MethodsandPropertiesIntroduced:
• IpfcTransform3D.Invert()
• IpfcTransform3D.TransformPoint()
• IpfcTransform3D.TransformVector()
WindowsandViews 183

• IpfcTransform3D.Matrix
• IpfcTransform3D.GetOrigin()
• IpfcTransform3D.GetXAxis()
• IpfcTransform3D.GetYAxis
• IpfcTransform3D.GetZAxis()
• CMpfcBase.MakeMatrixOrthonormal()
Allcoordinatesystems aretreatedin theVB API asif theywere three-
dimensional.Therefore,a pointin anyofthe coordinatesystems isalways
| representedby | theIpfcPoint3D               | class: |            |       |
| ------------- | ---------------------------- | ------ | ---------- | ----- |
| Vectorsstore  | thesamedatabutarerepresented |        | forclarity | bythe |
IpfcVector3Dclass.
Screencoordinatescontainaz-value whosepositive directionisoutwards from
thescreen.Thevalue ofzis notgenerallyimportantwhenspecifyinga screen
locationas an inputto amethod,butitisusefulin othersituations.For example,if
youselecta datumplane,you canfindthe directionofthe planeby calculatingthe
normaltothe plane,transformingtoscreen coordinates,thenlookingatthesign of
thez-coordinate.
Atransformationbetweentwocoordinatesystemsis representedby the
IpfcTransform3Dclass. Thisclass containsa 4x4matrixthatcombinesthe
conventional3x3matrix thatdescribestherelativeorientationof thetwosystems,
| andthevector     | thatdescribestheshift | betweenthem. |            |     |
| ---------------- | --------------------- | ------------ | ---------- | --- |
| The4x4matrixused | fortransformationsis  |              | asfollows: |     |
Theutility methodIpfcTransform3D.Invert() invertsa transformation
matrixsothatitcanbe usedto transformpoints inthe oppositedirection.
theVBAPI providestwoutilitiesfor performingcoordinatetransformations.The
| methodIpfcTransform3D.TransformPoint() |     |     |     | transforms athree- |
| -------------------------------------- | --- | --- | --- | ------------------ |
dimensionalpointandIpfcTransform3D.TransformVector()
| transformsathree-dimensional | vector. |     |     |     |
| ---------------------------- | ------- | --- | --- | --- |
Thefollowingdiagramsummarizesthe coordinatetransformations neededwhen
usingthe VBAPI andspecifiesthe theVB APImethodsthatprovidethe
transformationmatrix.
| 184 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

ThemethodCMpfcBase.MakeMatrixOrthonormal() convertsanon-
orthonormalmatrixto anorthonormal matrixwiththe specifiedscaling factor. The
| inputarguments        | follow:   |                |              |
| --------------------- | --------- | -------------- | ------------ |
| • Matrix—Thematrixto  |           | be convertedto | orthonormal. |
| • Scale—Scalefactorto |           | be appliedon   | thematrix.   |
| Transforming          | to Screen | Coordinates    |              |
MethodsandPropertiesIntroduced:
• IpfcView.Transform
• IpfcView.Rotate()
• IpfcViewOwner.GetCurrentViewTransform()
• IpfcViewOwner.SetCurrentViewTransform()
• IpfcViewOwner.CurrentViewRotate()
• IpfcTransform3D.Invert()
Theview matrixdescribesthe transformationfromsolidto screencoordinates.
ThepropertyIpfcView.Transform providestheview matrixfora modelin
theview.
WindowsandViews 185

Toget andsetthetransformationmatrix foramodelin thecurrentview,usethe
| methodsIpfcViewOwner.GetCurrentViewTransform() |     |     |     |     |     | and |     |
| ---------------------------------------------- | --- | --- | --- | --- | --- | --- | --- |
IpfcViewOwner.SetCurrentViewTransform().
ThemethodIpfcView.Rotate() rotatesan object,relativeto the X,Y, orZ
axisfor thespecifiedrotationangle.
| Torotatean | objectinthe | currentview, |     | usethemethod |     |     |     |
| ---------- | ----------- | ------------ | --- | ------------ | --- | --- | --- |
IpfcViewOwner.CurrentViewRotate().
Totransform fromscreen tosolidcoordinates, invertthe transformationmatrix
| usingthe     | methodIpfcTransform3D.Invert(). |            |     |        |       |             |     |
| ------------ | ------------------------------- | ---------- | --- | ------ | ----- | ----------- | --- |
| Transforming | to                              | Coordinate |     | System | Datum | Coordinates |     |
PropertyIntroduced:
• IpfcCoordSystem.CoordSys
| ThepropertyIpfcCoordSystem.CoordSys |     |     |     |     | providesthelocationand |     |     |
| ----------------------------------- | --- | --- | --- | --- | ---------------------- | --- | --- |
orientationofthecoordinatesystem datumin the coordinatesystem ofthesolid
thatcontains it. Thelocationisin termsof thedirectionsof thethreeaxes andthe
| positionofthe | origin. |     |             |     |     |     |     |
| ------------- | ------- | --- | ----------- | --- | --- | --- | --- |
| Transforming  | Window  |     | Coordinates |     |     |     |     |
PropertiesIntroduced
• IpfcWindow.ScreenTransform
• IpfcScreenTransform.PanX
• IpfcScreenTransform.PanY
• IpfcScreenTransform.Zoom
You canalterthe panandzoomof awindowby usinga ScreenTransformobject.
Thisobjectcontains threeattributes.PanXandPanY representthe horizontaland
verticalmovement. Everyincrementof1.0 movestheview pointonescreen width
orheight.Zoom representsa scalingfactorfor theview.This numbermustbe
greaterthanzero.
| Transforming | Coordinates |     |     | of an Assembly |     | Member |     |
| ------------ | ----------- | --- | --- | -------------- | --- | ------ | --- |
MethodIntroduced:
• IpfcComponentPath.GetTransform()
| ThemethodIpfcComponentPath.GetTransform() |     |     |     |     |     | providesthe | matrix |
| ----------------------------------------- | --- | --- | --- | --- | --- | ----------- | ------ |
fortransformingfrom thesolidcoordinatesystemofthe assemblymember to the
| solidcoordinatesofthe |     | parentassembly,or |     | thereverse. |                                 |     |     |
| --------------------- | --- | ----------------- | --- | ----------- | ------------------------------- | --- | --- |
| 186                   |     |                   |     |             | Creo®ParametricVBAPIUser’sGuide |     |     |

| ExampleCode-Normalizing          | aCoordinate | Transformation | Matrix   |
| -------------------------------- | ----------- | -------------- | -------- |
| thefilepfcViewExamples.vblocated |             |                | at<creo_ |
Thesamplecodein
| vbapi_loadpoint>/vbapi_examplesusestwomethodsto |         |              | transfer the |
| ----------------------------------------------- | ------- | ------------ | ------------ |
| viewtransformationfrom                          | oneview | to another.. |              |
WindowsandViews 187

14
ModelItem
SolidGeometryTraversal.........................................................................................189
GettingModelItemObjects.......................................................................................189
ModelItemInformation.............................................................................................190
DuplicatingModelItems............................................................................................191
LayerObjects..........................................................................................................191
Thischapterdescribes thetheVB APImethods thatenableyoutoaccess and
manipulateModelItems.
188 Creo®ParametricVBAPIUser’sGuide

Solid Geometry Traversal
Solidmodelsare madeup of11distinct typesof IpfcModelItem,as follows:
• IpfcFeature
• IpfcSurface
• IpfcEdge
• IpfcCurve(datumcurve)
• IpfcAxis(datumaxis)
• IpfcPoint(datumpoint)
• IpfcQuilt(datumquilt)
• IpfcLayer
• IpfcNote
• IpfcDimension
• IpfcRefDimension
Eachmodelitem isassigneda uniqueidentificationnumberthatwillnever
change.Inaddition, eachmodelitemcan beassigneda stringname.Layers,
points,axes,dimensions,andreferencedimensionsareautomatically assigneda
namethatcanbe changed.
Getting ModelItem Objects
MethodsandPropertiesIntroduced:
• IpfcModelItemOwner.ListItems()
• IpfcFeature.ListSubItems()
• IpfcLayer.ListItems()
• IpfcModelItemOwner.GetItemById()
• IpfcModelItemOwner.GetItemByName()
• IpfcFamColModelItem.RefItem
• IpfcSelection.SelItem
Allmodelsinheritfromthe interfaceIpfcModelItemOwner.The method
IpfcModelItemOwner.ListItems() returnsasequence of
IpfcModelItemscontainedin themodel. You canspecify whichtypeof
IpfcModelItemto collectbypassing inone oftheenumerated
EpfcModelItemTypevalues, oryoucancollectall IpfcModelItemsby
passingnullasthemodelitem type.
ModelItem 189

| Ifthemodelhasmultiplebodies,the |     |     | method              |     |
| ------------------------------- | --- | --- | ------------------- | --- |
| IpfcModelItemOwner.ListItems()  |     |     | returnstheexception |     |
IpfcXToolkitMultibodyUnsupported.
Note
The partmodelingfeaturesintroduced in CreoParametric1.0 willbe excluded
| from                           | thelist offeaturesreturnedby |     | themethod             |     |
| ------------------------------ | ---------------------------- | --- | --------------------- | --- |
| IpfcModelItemOwner.ListItems() |                              |     | if themodelitemtypeis |     |
specifiedasEpfcITEM_FEATURE. Forexampleeditround features,flexible
| modelingfeatures,andsoon             |     | willbe | excludedfrom | thelist. |
| ------------------------------------ | --- | ------ | ------------ | -------- |
| ThemethodsIpfcFeature.ListSubItems() |     |        |              | and      |
IpfcLayer.ListItems() producesimilarresultsfor specificfeaturesand
layers.These methodsreturna listofsubitems in thefeatureoritems in thelayer.
| Toaccess                              | specificmodelitems,callthemethod |             |     |                     |
| ------------------------------------- | -------------------------------- | ----------- | --- | ------------------- |
| IpfcModelItemOwner.GetItemById().This |                                  |             |     | methodsenablesyouto |
| accessthemodelitemby                  |                                  | identifier. |     |                     |
| Toaccess                              | specificmodelitems,callthemethod |             |     |                     |
IpfcModelItemOwner.GetItemByName().This
methodsenablesyouto
| accessthemodelitemby |     | name. |     |     |
| -------------------- | --- | ----- | --- | --- |
ThepropertyIpfcFamColModelItem.RefItem returnsthedimension or
| featureused                      | asaheader | for afamilytable. |         |                 |
| -------------------------------- | --------- | ----------------- | ------- | --------------- |
| ThepropertyIpfcSelection.SelItem |           |                   | returns | theitemselected |
| interactivelyby                  | theuser.  |                   |         |                 |
| ModelItem                        |           | Information       |         |                 |
MethodsandPropertiesIntroduced:
• IpfcModelItem.GetName()
• IpfcModelItem.SetName()
• IpfcModelItem.Id
• IpfcModelItem.Type
CertainIpfcModelItemsalso havea stringnamethatcanbe changedatany
| time.Themethods | IpfcModelItem.GetName()and |     |     |     |
| --------------- | -------------------------- | --- | --- | --- |
IpfcModelItem.SetName()accessthisname.
ThepropertyIpfcModelItem.Id returnstheunique integeridentifier forthe
IpfcModelItem.
| 190 |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

TheIpfcModelItem.Type propertyreturnsan enumerationobjectthat
indicatesthemodelitemtypeof thespecifiedIpfcModelItem.Seethe section
SolidGeometryTraversalonpage189for thelist ofpossiblemodelitemtypes.
| Duplicating | ModelItems |     |     |     |     |
| ----------- | ---------- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcBaseSession.AllowDuplicateModelItems()
You cancontrolthecreationofModelItemsmorethantwicefor thesameCreo
Parametricitem.Themethod
| IpfcBaseSession.AllowDuplicateModelItems() |     |     |     | allowsyouto | turn |
| ------------------------------------------ | --- | --- | --- | ----------- | ---- |
ONorOFF theoption toduplicate modelitems.Bydefault,thisoption isOFF.To
| turntheoption | ON,setthebooleanvalue |     | to FALSE. |     |     |
| ------------- | --------------------- | --- | --------- | --- | --- |
Note
Ifthis optionis nothandledproperlyon theapplicationside,itcancause
memorycorruption.Thus, althoughyoucan turnONandOFF thisoption as
| manytimes                        | asyouwant, | PTCrecommendsturningONand |           | OFFthisoption |     |
| -------------------------------- | ---------- | ------------------------- | --------- | ------------- | --- |
| only once,rightafterthesessionis |            |                           | obtained. |               |     |
| Layer                            | Objects    |                           |           |               |     |
IntheVB API,layersareinstances of IpfcModelItem.The followingsections
describehowtoget layerobjectsandtheoperations youcanperformonthem.
| Getting | Layer Objects |     |     |     |     |
| ------- | ------------- | --- | --- | --- | --- |
MethodIntroduced:
• IpfcModel.CreateLayer()
ThemethodIpfcModel.CreateLayer() returnsanewlayerwith thename
youspecify.
SeethesectionGettingModelItemObjectsonpage189for othermethodsthat
canreturnlayerobjects.
| Layer Operations |     |     |     |     |     |
| ---------------- | --- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcLayer.Status
| ModelItem |     |     |     |     | 191 |
| --------- | --- | --- | --- | --- | --- |

• IpfcLayer.ListItems()
• IpfcLayer.AddItem()
• IpfcLayer.RemoveItem()
• IpfcLayer.Delete()
• IpfcLayer.CountUnsupportedItems()
SupersededMethod:
• IpfcLayer.HasUnsupportedItems()
ThepropertyIpfcLayer.Status enablesyoutoaccess thedisplaystatusofa
layer.The correspondingenumerationclassis EpfcDisplayStatus andthe
| possiblevalues                       | areNormal, | Displayed,Blank,or |            | Hidden.              |        |
| ------------------------------------ | ---------- | ------------------ | ---------- | -------------------- | ------ |
| Usethe methodsIpfcLayer.ListItems(), |            |                    |            | IpfcLayer.AddItem(), |        |
| andIpfcLayer.RemoveItem()to          |            |                    | controlthe | contentsofa          | layer. |
Note
| Youcannotaddthefollowingitems |     |     | to alayer: |     |     |
| ----------------------------- | --- | --- | ---------- | --- | --- |
• ITEM_SURFACE,
• ITEM_EDGE,
• ITEM_COORD_SYS,
• ITEM_AXIS,
• ITEM_SIMPREP,
• ITEM_DTL_SYM_DEFINITION,
• ITEM_DTL_OLE_OBJECT,
• ITEM_EXPLODED_STATE.
| For theseitemsthe | methodwillthrowthe |     | exception |     |     |
| ----------------- | ------------------ | --- | --------- | --- | --- |
IpfcXToolkitInvalidType.
ThemethodIpfcLayer.Delete() removesthelayer(but nottheitems it
| contains)from | themodel. |     |     |     |     |
| ------------- | --------- | --- | --- | --- | --- |
ThemethodIpfcLayer.CountUnsupportedItems()
returnsthe number
ofitemtypesnotsupported asa IpfcModelItemobjectin thespecifiedlayer.
| Thismethoddeprecates | themethod |     |     |     |     |
| -------------------- | --------- | --- | --- | --- | --- |
IpfcLayer.HasUnsupportedItems().
| 192 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

15
Features
AccesstoFeatures..................................................................................................194
FeatureInformation.................................................................................................195
FeatureOperations.................................................................................................196
FeatureGroupsandPatterns...................................................................................198
UserDefinedFeatures.............................................................................................200
CreatingFeaturesfromUDFs...................................................................................201
AllCreoParametricsolidmodelsaremade upoffeatures.This chapterdescribes
howto programon thefeaturelevel usingthe VBAPI.
193

| Access to Features |     |     |     |
| ------------------ | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcFeature.ListChildren()
• IpfcFeature.ListParents()
• IpfcFeature.GroupDirectHeader
• IpfcFeatureGroup.GroupLeader
• IpfcFeaturePattern.PatternLeader
• IpfcFeaturePattern.ListMembers()
• IpfcSolid.ListFailedFeatures()
• IpfcSolid.ListFeaturesByType()
• IpfcSolid.GetFeatureById()
• IpfcBaseSession.QueryFeatureEdit()
| ThemethodsIpfcFeature.ListChildren() |     |     | and |
| ------------------------------------ | --- | --- | --- |
IpfcFeature.ListParents()returnasequence offeaturesthatcontainall
| thechildrenor parents | ofthespecifiedfeature. |     |     |
| --------------------- | ---------------------- | --- | --- |
ThepropertyIpfcFeature.GroupDirectHeader returnsthedirectheader
ofthegroup.
| Toget thefirstfeaturein | thespecifiedgroup | accessthe | property |
| ----------------------- | ----------------- | --------- | -------- |
IpfcFeatureGroup.GroupLeader.
| ThepropertyIpfcFeaturePattern.PatternLeader |     |     | and themethod |
| ------------------------------------------- | --- | --- | ------------- |
IpfcFeaturePattern.ListMembers() returnfeaturesthatmakeup the
specifiedfeaturepattern.Seethe sectionFeatureGroupsand Patternsonpage198
| formoreinformationon | featurepatterns. |     |     |
| -------------------- | ---------------- | --- | --- |
ThemethodIpfcSolid.ListFailedFeatures() returns asequencethat
containsallthe featuresthatfailedregeneration.
| ThemethodIpfcSolid.ListFeaturesByType() |     |     | returns asequenceof |
| --------------------------------------- | --- | --- | ------------------- |
featurescontainedin themodel. You canspecify whichtypeoffeaturetocollect
bypassingin oneof theEpfcFeatureType enumerationobjects,or youcan
collectall featuresbypassing voidnullas thetype. Ifyoulistall features,the
resultingsequence willinclude invisiblefeaturesthatCreo Parametriccreates
internally. Internalfeaturesareinvisiblefeatures usedinternallyforconstruction
purposes.Usethe method’s VisibleOnlyargumentto excludethem.Ifthe
argumentVisibleOnlyis True,the functionlists thepublicfeaturesonly.Ifthe
argumentis False,thefunction listsboth publicandinternalfeatures.
ThemethodIpfcSolid.GetFeatureById() returnsthe featureobjectwith
thecorrespondingintegeridentifier.
| 194 |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | ------------------------------- |

Afeaturecan be editedwiththe EditDefinitioncommand inCreo Parametric.The
methodIpfcBaseSession.QueryFeatureEdit(returns alist ofallthe
featuresthatare currentlybeing editedbythe EditDefinitioncommand.
Feature Information
PropertiesIntroduced:
• IpfcFeature.FeatType
• IpfcFeature.Status
• IpfcFeature.IsVisible
• IpfcFeature.IsReadonly
• IpfcFeature.IsEmbedded
• IpfcFeature.Number
• IpfcFeature.FeatTypeName
• IpfcFeature.FeatSubType
• IpfcRoundFeat.IsAutoRoundMember
TheenumerationclassesEpfcFeatureType andEpfcFeatureStatus
provideinformation fora specifiedfeature.The followingpropertiesspecify this
information:
• IpfcFeature.FeatType—Returnsthe typeof afeature.
• IpfcFeature.Status—Returnswhetherthefeatureis suppressed, active,
orfailedregeneration.
Theotherpropertiesthatgather featureinformation includethe following:
• IpfcFeature.IsVisible—Identifieswhetherthespecifiedfeaturewill
bevisible onthescreen.The methoddistinguishesvisiblefeaturesfrom
internalfeatures.Internal featuresareinvisiblefeaturesusedfor construction
purposes.
• IpfcFeature.IsReadonly—Identifies whetherthespecifiedfeaturecan
bemodified.
• IpfcFeature.IsEmbedded—Specifieswhetherthe specifiedfeatureis
anembedded datum.
• IpfcFeature.Number—Returnsthe featureregenerationnumber. This
methodreturnsvoidnullif thefeatureissuppressed.
ThepropertyIpfcFeature.FeatTypeName returnsa stringrepresentationof
thefeaturetype.
ThepropertyIpfcFeature.FeatSubType returns astringrepresentationof
thefeaturesubtype,for example,"Extrude"for aprotrusion feature.
Features 195

| ThepropertyIpfcRoundFeat.IsAutoRoundMember |            |         |                        | determineswhether |     |
| ------------------------------------------ | ---------- | ------- | ---------------------- | ----------------- | --- |
| thespecifiedroundfeatureis                 |            | amember | ofan AutoRoundfeature. |                   |     |
| Feature                                    | Operations |         |                        |                   |     |
MethodsandPropertiesIntroduced:
• IpfcSolid.ExecuteFeatureOps()
• IpfcFeature.CreateSuppressOp()
• IpfcSuppressOperation.Clip
• IpfcSuppressOperation.AllowGroupMembers
• IpfcSuppressOperationAllowChildGroupMembers
• IpfcFeature.CreateDeleteOp
• IpfcDeleteOperation.Clip
• IpfcDeleteOperation.AllowGroupMembers
• IpfcDeleteOperation.AllowChildGroupMembers
• IpfcDeleteOperation.KeepEmbeddedDatums
• IpfcFeatureCreateResumeOp
• IpfcResumeOperation.WithParents
• IpfcFeature.CreateReorderBeforeOp
• IpfcReorderBeforeOperation.BeforeFeat
• IpfcFeature.CreateReorderAfterOp
• IpfcReorderAfterOperation.AfterFeat
| ThemethodIpfcSolid.ExecuteFeatureOps() |     |     |     | causesasequence | of  |
| -------------------------------------- | --- | --- | --- | --------------- | --- |
featureoperations torunin order.Featureoperationsinclude suppressing,
| resuming,reordering,and |                      | deletingfeatures.The | optional                      |     |     |
| ----------------------- | -------------------- | -------------------- | ----------------------------- | --- | --- |
| IpfcRegenInstructions   |                      | argument             | specifieswhethertheuserwillbe |     |     |
| allowedtofixthe         | modelifaregeneration |                      | failureoccurs.                |     |     |
Note
Regeneratingmodelsin resolve_modeisnotsupported.As aresult,the
methodIpfcSolid.ExecuteFeatureOps()
|     |     |     |     | isnotsupported | anymore. |
| --- | --- | --- | --- | -------------- | -------- |
You cancreatean operationthatwilldelete,suppress, reorder, orresume certain
featuresusingthe methodsin theclassIpfcFeature.Eachcreated operation
mustbe passedasa memberof theIpfcFeatureOperations objectto the
methodIpfcSolid.ExecuteFeatureOps().
| 196 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

Someoftheoperationshavespecific optionsthatyoucan modifytocontrolthe
| behaviorofthe | operation: |     |     |     |     |
| ------------- | ---------- | --- | --- | --- | --- |
• Clip—Specifies whetherto deleteor suppressallfeaturesafterthe selected
| feature.By                 | default,this                          | optionis false. |              |         |     |
| -------------------------- | ------------------------------------- | --------------- | ------------ | ------- | --- |
| Usethe                     | propertiesIpfcDeleteOperation.Clipand |                 |              |         |     |
| IpfcSuppressOperation.Clip |                                       |                 | tomodifythis | option. |     |
• AllowGroupMembers—Ifthis optionis settotrue andif thefeatureto be
deletedor suppressedisamember ofa group,thenthefeaturewillbe deleted
orsuppressedoutof thegroup. Ifthisoption issetto false,thentheentire
groupcontainingthe featureisdeleted orsuppressed. Bydefault,this option is
| false.It                                 | canbe setto                                       | trueonly ifthe | optionClipis | settotrue. |               |
| ---------------------------------------- | ------------------------------------------------- | -------------- | ------------ | ---------- | ------------- |
| Usethe                                   | propertiesIpfcSuppressOperation.AllowGroupMembers |                |              |            |               |
| andIpfcDeleteOperation.AllowGroupMembers |                                                   |                |              |            | to modifythis |
option.
| • AllowChildGroupMembers—If |     |     | thisoption | issetto true | andifthe |
| --------------------------- | --- | --- | ---------- | ------------ | -------- |
childrenof thefeatureto be deletedorsuppressed aremembersof agroup,
thenthechildrenof thefeaturewillbe individuallydeletedor suppressedout
ofthegroup. Ifthisoption issetto false, thenthe entiregroupcontainingthe
featureanditschildrenis deletedor suppressed. Bydefault,this optionis
| false.It          | canbe setto | trueonly ifthe | optionsClipand |     |     |
| ----------------- | ----------- | -------------- | -------------- | --- | --- |
| AllowGroupMembers |             | aresetto       | true.          |     |     |
| Usethe            | properties  |                |                |     |     |
IpfcSuppressOperation.AllowChildGroupMembersand
IpfcDeleteOperation.AllowChildGroupMembers
to modifythis
option.
• KeepEmbeddedDatums—Specifieswhetherto retaintheembeddeddatums
storedin afeaturewhiledeletingthefeature.Bydefault,thisoption isfalse.
| Usethe | propertyIpfcDeleteOperation.KeepEmbeddedDatums |     |     |     | to  |
| ------ | ---------------------------------------------- | --- | --- | --- | --- |
modifythisoption.
• WithParents—Specifieswhetherto resume theparents oftheselected
feature.
| Usethe | propertyIpfcResumeOperation.WithParentsto |     |     |     | modifythis |
| ------ | ----------------------------------------- | --- | --- | --- | ---------- |
option.
• BeforeFeat—Specifiesthefeaturebeforewhichyouwant to reorderthe
features.
| Usethe | propertyIpfcReorderBeforeOperation.BeforeFeat |     |     |     | to  |
| ------ | --------------------------------------------- | --- | --- | --- | --- |
modifythisoption.
Features 197

• AfterFeat—Specifiesthe featureafterwhichyouwantto reorderthe
features.
| Usethe | propertyIpfcReorderAfterOperation.AfterFeat |     |     |     |     | to  |
| ------ | ------------------------------------------- | --- | --- | --- | --- | --- |
modifythisoption.
| Feature | Groups |     | and | Patterns |     |     |
| ------- | ------ | --- | --- | -------- | --- | --- |
Patternsaretreatedasfeaturesin Creo Parametric.Afeaturetype,FEATTYPE_
| PATTERN_HEAD,is |     | usedfor | thepatternheader |     | feature. |     |
| --------------- | --- | ------- | ---------------- | --- | -------- | --- |
Note
The patternheaderfeatureis nottreatedas aleaderoramember ofthe pattern
| by themethodsdescribed |     |     | in thefollowingsection. |     |     |     |
| ---------------------- | --- | --- | ----------------------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcFeature.Group
• IpfcFeature.Pattern
• IpfcSolid.CreateLocalGroup
• IpfcFeatureGroup.Pattern
• IpfcFeatureGroup.GroupLeader()
• IpfcFeaturePattern.PatternLeader
• IpfcFeaturePattern.ListMembers()
• IpfcFeaturePattern.Delete()
ThepropertyIpfcFeature.Group returnsa handletothe localgroup that
containsthespecifiedfeature.
| Toget thefirstfeaturein |     |     | thespecifiedgroup |     | calltheproperty |     |
| ----------------------- | --- | --- | ----------------- | --- | --------------- | --- |
IpfcFeatureGroup.GroupLeader.
| ThepropertyIpfcFeaturePattern.PatternLeader |     |     |     |     |     | and themethod |
| ------------------------------------------- | --- | --- | --- | --- | --- | ------------- |
IpfcFeaturePattern.ListMembers() returnfeaturesthatmakeup the
specifiedfeaturepattern.
Apatterniscomposedofa patternheaderfeatureanda numberofmember
features.Youcan patternonly asinglefeature.Topatternseveralfeatures,createa
| localgroup | andpatternthis |     | group. |     |     |     |
| ---------- | -------------- | --- | ------ | --- | --- | --- |
You canalsocreatea patternofpattern.This createsamultiple levelpattern.From
| CreoParametric2.0M170onward,for  |     |     |        | apatternof | pattern,themethod               |               |
| -------------------------------- | --- | --- | ------ | ---------- | ------------------------------- | ------------- |
| IpfcFeaturePattern.ListMembers() |     |     |        |            | returnsallthe                   | patternheader |
| featurescreatedatthefirst        |     |     | level. |            |                                 |               |
| 198                              |     |     |        |            | Creo®ParametricVBAPIUser’sGuide |               |

Forexample,consideramodelwhereapatternof patternhasbeencreated.The
| modeltreeisas shownbelow: |     |     |
| ------------------------- | --- | --- |
ThemethodIpfcFeaturePattern.ListMembers() returns thepattern
| headerfeatureswith | followingIDsfora | patternofpattern: |
| ------------------ | ---------------- | ----------------- |
• 119
• 177
• 221
• 265
| ThepropertiesIpfcFeature.Pattern |     | and                         |
| -------------------------------- | --- | --------------------------- |
| IpfcFeatureGroup.Patternreturn   |     | theIpfcFeaturePatternobject |
thatcontains thecorrespondingIpfcFeature orIpfcFeatureGroup. Use
themethodIpfcSolid.CreateLocalGroup()
to takeasequence of
Features 199

| featuresandcreatealocalgroup |     | with | thespecifiedname.Todeletea |     |
| ---------------------------- | --- | ---- | -------------------------- | --- |
IpfcFeaturePatternobject,callthemethod
IpfcFeaturePattern.Delete().
| Notes On | Feature | Groups |     |     |
| -------- | ------- | ------ | --- | --- |
Featuregroupshaveagroup header feature,whichshowsup in themodel
informationandfeaturelist forthemodel. This featurewillbe insertedin the
| regenerationlist | toa positionjustbefore |                      | thefirstfeaturein | thegroup. |
| ---------------- | ---------------------- | -------------------- | ----------------- | --------- |
| Theresultsof     | theheader              | featureareasfollows: |                   |           |
• Modelsthatcontaingroups willgetone extrafeatureinthe regenerationlist,
oftypeEpfcFEATTYPE_GROUP_HEAD. Thisaffects thefeaturenumbersof
| allsubsequentfeatures,includingthosein |     |     | thegroup. |     |
| -------------------------------------- | --- | --- | --------- | --- |
• Eachgroup automaticallycontains theheaderfeaturein the listoffeatures
| returnedfrom | IpfcFeatureGroup.ListMembers(). |     |     |     |
| ------------ | ------------------------------- | --- | --- | --- |
• Eachgroup automaticallygetsthe group headfeatureas theleader.This is
| returnedfrom | IpfcFeatureGroup.GroupLeader. |     |     |     |
| ------------ | ----------------------------- | --- | --- | --- |
• Eachgroup patterncontainsa seriesof groups,and eachgroup in thepattern
| willbesimilarly | constructed. |          |     |     |
| --------------- | ------------ | -------- | --- | --- |
| User            | Defined      | Features |     |     |
Groupsin CreoParametric representsetsofcontiguous featuresthatactasa
singlefeaturefor specificoperations.Individualfeaturesareaffected bymost
| operationswhilesomeoperationsapplyto |     |     | an entiregroup: |     |
| ------------------------------------ | --- | --- | --------------- | --- |
• Suppress
• Delete
• Layers
• Patterning
UserdefinedFeatures (UDFs)aregroups offeaturesthatarestoredin afile.When
aUDFis placedina newmodelthecreatedfeaturesareautomatically assignedto
agroup.Alocal groupis asetoffeaturesthathavebeenspecificallyassignedto a
| groupto makemodifications |     | andpatterningeasier. |     |     |
| ------------------------- | --- | -------------------- | --- | --- |
Note
| Allmethodsin | this | sectioncanbe | usedfor UDFsandlocal            | groups. |
| ------------ | ---- | ------------ | ------------------------------- | ------- |
| 200          |      |              | Creo®ParametricVBAPIUser’sGuide |         |

| Read | Access |     | to Groups |     |     | and | User Defined | Features |
| ---- | ------ | --- | --------- | --- | --- | --- | ------------ | -------- |
MethodsandPropertiesIntroduced:
• IpfcFeatureGroup.UDFName
• IpfcFeatureGroup.UDFInstanceName
• IpfcFeatureGroup.ListUDFDimensions()
• IpfcUDFDimension.GetUDFDimensionName
Userdefinedfeatures(UDF’s)aregroups offeaturesthatcanbe storedin afile
andaddedto anewmodel. Alocalgroup issimilar toa UDFexceptitisavailable
| onlyin themodelin |     |     | whichiswas |     | created. |     |     |     |
| ----------------- | --- | --- | ---------- | --- | -------- | --- | --- | --- |
ThepropertyIpfcFeatureGroup.UDFName providesthename ofthegroup
forthespecifiedgroup instance.Aparticulargroupdefinitioncanbe usedmore
| thanoncein                                    | aparticular |     | model.            |     |     |          |                     |              |
| --------------------------------------------- | ----------- | --- | ----------------- | --- | --- | -------- | ------------------- | ------------ |
| Ifthegroup                                    | isafamily   |     | tableinstance,the |     |     | property |                     |              |
| IpfcFeatureGroup.UDFInstanceName              |             |     |                   |     |     |          | suppliestheinstance | name.        |
| ThemethodIpfcFeatureGroup.ListUDFDimensions() |             |     |                   |     |     |          |                     | traversesthe |
dimensionsthatbelongto the UDF.These dimensionscorrespond tothe
dimensionsspecifiedasvariableswhenthe UDFwas created.Dimensionsof the
originalfeaturesthatwere notvariablesinthe UDFarenot includedunlessthe
| UDFwas                                          | placedusingthe |     |     | Independentoption. |     |     |     |          |
| ----------------------------------------------- | -------------- | --- | --- | ------------------ | --- | --- | --- | -------- |
| ThepropertyIpfcUDFDimension.GetUDFDimensionName |                |     |     |                    |     |     |     | provides |
accesstothe dimensionname specifiedwhen theUDF wascreated,and notthe
nameofthe dimensionin thecurrentmodel.This nameis requiredto placethe
UDFprogrammaticallyusingthe methodIpfcSolid.CreateUDFGroup().
| Creating |     | Features |     |     | from |     | UDFs |     |
| -------- | --- | -------- | --- | --- | ---- | --- | ---- | --- |
MethodIntroduced:
• IpfcSolid.CreateUDFGroup()
ThemethodIpfcSolid.CreateUDFGroup() is usedto createnew features
byretrieving andapplyingthe contentsof anexisting UDFfile. Itis equivalentto
| theCreoParametriccommand |     |     |     | Feature,Create, |     |     | UserDefined. |     |
| ------------------------ | --- | --- | --- | --------------- | --- | --- | ------------ | --- |
Tounderstandthe followingexplanation ofthis method,youmusthavea good
knowledgeand understandingof theuseofUDF’s in CreoParametric.PTC
recommendsthatyoureadaboutUDF’sin theCreo Parametrichelp, andpractice
definingandusing UDF’s inCreo Parametricbeforeyouattempt tousethis
method.
Features 201

WhenyoucreateaUDF interactively, CreoParametricpromptsyoufor the
informationitneedstofixthe propertiesoftheresulting features.Whenyou
createa UDFfrom theVB API,youcan providesomeor allofthis information
programmaticallyby fillingseveralcompactdataclassesthatareinputsto the
methodIpfcSolid.CreateUDFGroup.
| Duringthecallto | IpfcSolid.CreateUDFGroup(),Creo |     |     |     |     |     | Parametric |
| --------------- | ------------------------------- | --- | --- | --- | --- | --- | ---------- |
| promptsyoufor   | thefollowing:                   |     |     |     |     |     |            |
• Informationrequiredby theUDF thatwas notprovided inthe inputdata
structures.
| • Correctinformationto |     |     | replaceerroneousinformation |     |     |     |     |
| ---------------------- | --- | --- | --------------------------- | --- | --- | --- | --- |
Suchpromptsareausefulwayof diagnosingerrors whenyoudevelopyour
application.Thisalso meansthat, in additiontocreatingUDF’s programmatically
| toprovideautomaticsynthesisof |     |     |     | modelgeometry,youcanalso |     |     | use |
| ----------------------------- | --- | --- | --- | ------------------------ | --- | --- | --- |
pfcSolid.CreateUDFGroup()to createUDF’ssemi-interactively.This can
simplifythe interactionsneededto placeacomplexUDF makingiteasierfor the
userandlesspronetoerror.
| Creating                                      | UDFs |     |           |     |             |             |        |
| --------------------------------------------- | ---- | --- | --------- | --- | ----------- | ----------- | ------ |
| Creatinga UDFrequiresthefollowinginformation: |      |     |           |     |             |             |        |
| • Name—Thenameofthe                           |      |     | UDFyouare |     | creatingand | theinstance | nameif |
applicable.
• Dependency—Specifyif theUDF isindependent oftheUDF definitionor is
| modifiedbythe |     | changersmadeto |     | it.        |                    |     |     |
| ------------- | --- | -------------- | --- | ---------- | ------------------ | --- | --- |
| • Scale—Howto |     | scaletheUDF    |     | relativeto | theplacementmodel. |     |     |
• VariableDimension—Thenewvalues ofthe variablesdimensions andpattern
parameters,thosewhose valuescanbe modifiedeachtime theUDFis created.
• DimensionDisplay—Whether toshowor blanknon-variabledimensions
| createdwithin | theUDF |     | group. |     |     |     |     |
| ------------- | ------ | --- | ------ | --- | --- | --- | --- |
• References—ThegeometricalelementsthattheUDF needsinorder to relate
| thefeaturesitcontainsto |     |     | the | existingmodelsfeatures.Theelements |     |     |     |
| ----------------------- | --- | --- | --- | ---------------------------------- | --- | --- | --- |
correspondto thepicksthatCreo Parametricpromptsyouforwhen youcreate
aUDF interactivelyusing thepromptsdefinedwhen theUDF wascreated.
| You cannotselectan |     |     | embeddeddatumas |     | theUDF | reference. |     |
| ------------------ | --- | --- | --------------- | --- | ------ | ---------- | --- |
• PartsIntersection—When aUDF thatisbeingcreated inan assembly contains
featuresthatmodifytheexistinggeometryyoumustdefine whichparts are
affectedorintersected.Youalso needtoknow atwhatlevel inan assembly
| eachintersectionis |     | goingtobe |     | visible. |     |                                 |     |
| ------------------ | --- | --------- | --- | -------- | --- | ------------------------------- | --- |
| 202                |     |           |     |          |     | Creo®ParametricVBAPIUser’sGuide |     |

• Orientations—Whena UDFcontains afeaturewith adirectionthatisdefined
inrespectto adatumplane CreoParametricmustknowwhat directionthe new
| featurewillpoint           |     | to.Whenyoucreatesuch |                 |     | aUDFinteractivelyCreo |             |
| -------------------------- | --- | -------------------- | --------------- | --- | --------------------- | ----------- |
| Parametricpromptyouforthis |     |                      | informationwith |     |                       | afliparrow. |
• Quadrants—WhenaUDFcontains alinearlyplacedfeaturethatreferences
twodatumplanesto defineit’s locationinthe newmodelCreo Parametric
promptsyoutopickthe locationofthenewfeature.This isdeterminedby
whichsideof eachdatumplane thefeaturemustlie.This selectionisreferred
toas thequadrant becausethearefourpossiblecombinationsfor eachlinearly
placefeature.
Topassall theabovevalues toCreo Parametric,theVB APIusesaspecial class
thatpreparesand setsalltheoptions andpasses themto Creo Parametric.
| Creating | Interactively |     | Defined |     | UDFs |     |
| -------- | ------------- | --- | ------- | --- | ---- | --- |
MethodIntroduced:
• CCpfcUDFPromptCreateInstructions.Create()
Thisstaticmethodisused tocreatean instructionsobjectthatcanbe usedto
prompta userforthe requiredvaluesthatwillcreateaUDF interactively.
| Creating | a Custom |     | UDF |     |     |     |
| -------- | -------- | --- | --- | --- | --- | --- |
MethodIntroduced:
• CCpfcUDFCustomCreateInstructions.Create()
Thismethodcreatesa IpfcUDFCustomCreateInstructionsobjectwith a
specifiedname.Tosetthe UDFcreationparametersprogrammaticallyyoumust
modifythis objectasdescribed below.The members ofthisclass relatecloselyto
thepromptsCreoParametricgivesyouwhen youcreateaUDF interactively. PTC
recommendsthatyouexperimentwithcreatingthe UDFinteractivelyusingCreo
| ParametricbeforeyouwritethetheVB |            |       |          | API | codeto | fillthestructure. |
| -------------------------------- | ---------- | ----- | -------- | --- | ------ | ----------------- |
| Setting                          | the Family | Table | Instance |     | Name   |                   |
PropertyIntroduced:
• IpfcUDFCustomCreateInstructions.InstanceName
IftheUDF containsafamily table,thisfieldcan beused toselectthe instancein
thetable.Ifthe UDFdoesnot containafamilytable, orif thegenericinstance is
| tobe selected,the |     | donot setthestring. |     |     |     |     |
| ----------------- | --- | ------------------- | --- | --- | --- | --- |
Features 203

Setting Dependency Type
PropertyIntroduced:
• IpfcUDFCustomCreateInstructions.DependencyType
TheEpfcUDFDependencyType objectrepresents thedependencytypeofthe
UDF.The choicescorrespond tothe choicesavailablewhen youcreateaUDF
interactively. Thisenumerated typetakesthe followingvalues:
• EpfcUDFDEP_INDEPENDENT
• EpfcUDFDEP_DRIVEN
Note
EpfcUDFDEP_INDEPENDENT isthedefault value,ifthis optionis notset.
Setting Scale and Scale Type
PropertiesIntroduced:
• IpfcUDFCustomCreateInstructions.ScaleType
• IpfcUDFCustomCreateInstructions.Scale
ThepropertyIpfcUDFCustomCreateInstructions.ScaleType
specifiesthelengthunitsofthe UDFin theform ofthe EpfcUDFScaleType
object.This enumeratedtypetakesthefollowingvalues:
• EpfcUDFSCALE_SAME_SIZE
• EpfcUDFSCALE_SAME_DIMS
• EpfcUDFSCALE_CUSTOM
• EpfcUDFSCALE_nil
Note
The defaultvalueisEpfcUDFSCALE_SAME_SIZE if thisoption isnotset.
ThepropertyIpfcUDFCustomCreateInstructions.Scale specifies the
scalefactor. IftheScaleTypeis settoEpfcUDFSCALE_CUSTOM, theproperty
IpfcUDFCustomCreateInstructions.Scaleassignsthe userdefined
scalefactor. Otherwise,thisattribute isignored.
204 Creo®ParametricVBAPIUser’sGuide

| Setting | the Appearance |     |     | of  | the Non | UDF | Dimensions |     |     |
| ------- | -------------- | --- | --- | --- | ------- | --- | ---------- | --- | --- |
PropertyIntroduced:
• IpfcUDFCustomCreateInstructions.DimDisplayType
| TheobjectEpfcUDFDimensionDisplayTypesetsthe |     |     |     |     |     |     | options |     | inCreo |
| ------------------------------------------- | --- | --- | --- | --- | --- | --- | ------- | --- | ------ |
Parametricfordeterminingtheappearancein themodelofUDF dimensionsand
patternparametersthatwere notvariableinthe UDF, andtherefore cannotbe
modifiedinthe model.This enumeratedtypetakesthefollowingvalues:
• EpfcUDFDISPLAY_NORMAL
• EpfcUDFDISPLAY_READ_ONLY
• EpfcUDFDISPLAY_BLANK
Note
defaultvalueisEpfcUDFDISPLAY_NORMAL
| The     |              |     |            |     |     |     | ifthis     | optionis | notset. |
| ------- | ------------ | --- | ---------- | --- | --- | --- | ---------- | -------- | ------- |
| Setting | the Variable |     | Dimensions |     |     | and | Parameters |          |         |
MethodsandPropertiesIntroduced:
• IpfcUDFCustomCreateInstructions.VariantValues
• CCpfcUDFVariantDimension.Create()
• CCpfcUDFVariantPatternParam.Create()
IpfcUDFVariantValuesclass representsan arrayofvariabledimensionsand
patternparameters.
| CCpfcUDFVariantDimension.Create()is |     |     |     |     |            |                      | astaticmethodcreatinga |     |     |
| ----------------------------------- | --- | --- | --- | --- | ---------- | -------------------- | ---------------------- | --- | --- |
| IpfcUDFVariantDimension.It          |     |     |     |     | acceptsthe | followingparameters: |                        |     |     |
• Name—Thesymbol thatthedimension hadwhenthe UDFwas originally
definednotthepromptthattheUDF useswhenitis createdinteractively.To
makethis nameeasyto remember, beforeyoudefinethe UDFthatyouplanto
createwiththe VBAPI,you shouldmodifythesymbols ofall thedimensions
| thatyouwantto              |     | selectto |     | bevariable. | Ifyougetthe |     | namewrong,               |     |     |
| -------------------------- | --- | -------- | --- | ----------- | ----------- | --- | ------------------------ | --- | --- |
| IpfcSolid.CreateUDFGroup() |     |          |     |             | willnot     |     | recognizethedimensionand |     |     |
promptstheuserfor thevaluein theusualwaydoes notmodifythevalue.
| • DimensionValue—The                 |       |                 | newvalue.      |     |          |     |               |     |       |
| ------------------------------------ | ----- | --------------- | -------------- | --- | -------- | --- | ------------- | --- | ----- |
| Ifyoudonotremember                   |       |                 | thename,youcan |     | finditby |     | creatingthe   | UDF |       |
| interactivelyin                      | atest | model,thenusing |                |     | the      |     |               |     |       |
| IpfcFeatureGroup.ListUDFDimensions   |       |                 |                |     |          |     | and           |     |       |
| IpfcUDFDimension.GetUDFDimensionName |       |                 |                |     |          |     | to findoutthe |     | name. |
| Features                             |       |                 |                |     |          |     |               |     | 205   |

| CCpfcUDFVariantPatternParam.Create() |                               |     |     |     | isa staticmethod    | which |
| ------------------------------------ | ----------------------------- | --- | --- | --- | ------------------- | ----- |
| createsa                             | IpfcUDFVariantPatternParam.It |     |     |     | acceptsthefollowing |       |
parameters:
• name—Thestringnamethatthe patternparameter hadwhenthe UDFwas
originallydefined
• patternparam—Thenewvalue.
| AftertheIpfcUDFVariantValues                  |          |                    |            | objecthasbeencompiled,use |       |     |
| --------------------------------------------- | -------- | ------------------ | ---------- | ------------------------- | ----- | --- |
| IpfcUDFCustomCreateInstructions.VariantValues |          |                    |            |                           | toadd | the |
| variabledimensions                            |          | andparameterstothe |            | instructions.             |       |     |
| Setting                                       | the User | Defined            | References |                           |       |     |
MethodsandPropertiesIntroduced:
• CCpfcUDFReference.Create()
• IpfcUDFReference.IsExternal
• IpfcUDFReference.ReferenceItem
• IpfcUDFCustomCreateInstructions.References
ThemethodCCpfcUDFReference.Create() is astaticmethodcreatinga
IpfcUDFReferenceobject.Itacceptsthe
followingparameters:
• PromptForReference—Thepromptdefinedforthis referencewhentheUDF
wasoriginallysetup.It indicateswhichreferencethisstructure isproviding. If
| yougetthepromptwrong,           |     |     | pfcSolid.CreateUDFGroup()willnot |                |              |     |
| ------------------------------- | --- | --- | -------------------------------- | -------------- | ------------ | --- |
| recognizeitandpromptstheuserfor |     |     |                                  | thereferencein | theusualway. |     |
• ReferenceItem—Specifiesthe IpfcSelection objectrepresentingthe
referencedelement.You cansetSelection programmaticallyorpromptthe
userfora selectionseparately. You cannotsetan embeddeddatumastheUDF
reference.
Therearetwotypesofreference:
○ Internal—Thereferencedelement belongsdirectlyto themodelthatwill
containthe UDF. Foran assembly,thismeans thattheelementbelongs to
thetoplevel.
○ External—Thereferencedelementbelongs to anassembly member other
thantheplacementmember.
| Tosetthe | referencetype, | usetheproperty |     |     |     |     |
| -------- | -------------- | -------------- | --- | --- | --- | --- |
IpfcUDFReference.IsExternal.
| Tosetthe | itemto | be usedforreference, |     | usetheproperty |     |     |
| -------- | ------ | -------------------- | --- | -------------- | --- | --- |
IpfcUDFReference.ReferenceItem.
| 206 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

| AftertheIpfcUDFReferences                  |     | objecthasbeenset, | use       |     |
| ------------------------------------------ | --- | ----------------- | --------- | --- |
| IpfcUDFCustomCreateInstructions.References |     |                   | to addthe |     |
program-definedreferences.
| Setting | the Assembly | Intersections |     |     |
| ------- | ------------ | ------------- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcUDFAssemblyIntersection.Create()
• IpfcUDFAssemblyIntersection.InstanceNames
• IpfcUDFCustomCreateInstructions.Intersections
CCpfcUDFAssemblyIntersection.Create()is astaticmethodcreating
| aIpfcUDFReferenceobject.Itaccepts |     | thefollowingparameters: |     |     |
| --------------------------------- | --- | ----------------------- | --- | --- |
• ComponentPath—Isan intseqtypeobjectrepresentingthecomponentpath
| ofthepart | to be intersected. |     |     |     |
| --------- | ------------------ | --- | --- | --- |
• Visibility level—Thenumberthatcorrespondsto the visibilitylevel ofthe
intersectedpartin theassembly.If thenumberisequal tothe lengthofthe
componentpaththe featureisvisiblein thepart thatitintersects.IfVisibility
levelis 0,thefeatureisvisibleatthe levelof theassemblycontainingthe
UDF.
IpfcUDFAssemblyIntersection.InstanceNames setsanarrayof
namesfor thenewinstances ofparts createdto representthe intersection
| geometry.This       | propertyacceptsthefollowingparameters: |     |     |     |
| ------------------- | -------------------------------------- | --- | --- | --- |
| • instancenames—isa | com.ptc.cipjava.stringseqtypeobject    |     |     |     |
representingthearrayofnewinstance names.
| AftertheIpfcUDFAssemblyIntersections          |     |     | objecthasbeenset, | use       |
| --------------------------------------------- | --- | --- | ----------------- | --------- |
| IpfcUDFCustomCreateInstructions.Intersections |     |     |                   | toadd the |
assemblyintersections.
| Setting | Orientations |     |     |     |
| ------- | ------------ | --- | --- | --- |
PropertiesIntroduced:
• IpfcUDFCustomCreateInstructions.Orientations
IpfcUDFOrientations classrepresents anarrayof orientationsthatprovide
theanswersto CreoParametricpromptsthatuseaflip arrow. Eachtermisa
EpfcUDFOrientationobjectthattakesthefollowingvalues:
• EpfcUDFORIENT_INTERACTIVE—Prompt for theorientationusing aflip
arrow.
EpfcUDFORIENT_NO_FLIP—Acceptthedefaultflip
• orientation.
| Features |     |     |     | 207 |
| -------- | --- | --- | --- | --- |

| • EpfcUDFORIENT_FLIP—Invert |     |     | theorientationfrom | thedefault |
| --------------------------- | --- | --- | ------------------ | ---------- |
orientation.
Theorder oforientationsshould correspondto theorder in whichCreoParametric
promptsfor themwhen theUDF iscreated interactively. Ifyoudonot providean
orientationthatCreoParametricneeds, itusesthedefaultvalue NO_FLIP.
AftertheIpfcUDFOrientations
objecthasbeensetuse
| IpfcUDFCustomCreateInstructions.Orientation |     |     |     | toadd the |
| ------------------------------------------- | --- | --- | --- | --------- |
orientations.
| Setting | Quadrants |     |     |     |
| ------- | --------- | --- | --- | --- |
PropertyIntroduced:
• IpfcUDFCustomCreateInstructions.Quadrants
ThepropertyIpfcUDFCustomCreateInstructions.Quadrants setsan
arrayof points,whichprovidetheX, Y,andZ coordinatesthatcorrespond to the
picksanswering theCreo Parametricpromptsforthe featurepositions.Theorder
ofquadrantsshould correspondto theorder in whichCreoParametricpromptsfor
| themwhenthe | UDFis createdinteractively. |            |     |     |
| ----------- | --------------------------- | ---------- | --- | --- |
| Setting     | the External                | References |     |     |
PropertyIntroduced:
• IpfcUDFCustomCreateInstructions.ExtReferences
ThepropertyIpfcUDFCustomCreateInstructions.ExtReferences
setsanexternalreferenceassemblyto be usedwhen placingtheUDF.This willbe
requiredwhenplacing theUDF inthe componentusing referencesoutsideofthat
component.Referencescouldbe to thetoplevel assembly ofanothercomponent.
| Example | Code 1 |     |     |     |
| ------- | ------ | --- | --- | --- |
Thesamplecodein thefilepfcFeaturesExamples.vblocated at<creo_
| vbapi_loadpoint>/vbapi_examplescopies |     |     | ofanode | UDFata |
| ------------------------------------- | --- | --- | ------- | ------ |
particularcoordinatesystemlocationin apart. Thenode UDFis asphericalcut
centeredatthe coordinatesystem whosediameter isdrivenby the'diam' argument
tothe method..
| Example | Code 2 |     |     |     |
| ------- | ------ | --- | --- | --- |
Thesamplecodein thefilepfcFeaturesExamples.vblocated at<creo_
| vbapi_loadpoint>/vbapi_examplesplaces |     |     | copiesof | aholeUDFata |
| ------------------------------------- | --- | --- | -------- | ----------- |
particularlocationin an assembly. Theholeisembedded ina surfaceof oneofthe
assembly'scomponents, andplacedaparticular locationawayfrom twonormal
datumplanes (thedefault valueforthe dimensionis usedforthis example).The
| 208 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

UDFcreationrequiresaquadrant determiningthe locationfortheUDF (sinceit
couldbeone offourareas)andintersection instructionsfor theassembly members
(thisexamplemakes theholevisibledownto thepart level).The methodreturns
theIpfcFeatureGroupobjectcreated.
Features 209

16
Datum Features
DatumPlaneFeatures.............................................................................................211
DatumAxisFeatures...............................................................................................213
GeneralDatumPointFeatures.................................................................................214
DatumCoordinateSystemFeatures.........................................................................215
Thischapterdescribes theVB APImethodsand propertiesthatprovideread
accesstothe propertiesofdatumfeatures.
210 Creo®ParametricVBAPIUser’sGuide

| Datum              | Plane | Features   |                        |     |     |     |
| ------------------ | ----- | ---------- | ---------------------- | --- | --- | --- |
| Thepropertiesofthe |       | DatumPlane | featurearedefinedinthe |     |     |     |
IpfcDatumPlaneFeatdataobject.
MethodsandPropertiesIntroduced:
• IpfcDatumPlaneFeat.Flip
• IpfcDatumPlaneFeat.Constraints
• IpfcDatumPlaneConstraint.ConstraintType
• IpfcDatumPlaneThroughConstraint.ThroughRef
• IpfcDatumPlaneNormalConstraint.NormalRef
• IpfcDatumPlaneParallelConstraint.ParallelRef
• IpfcDatumPlaneTangentConstraint.TangentRef
• IpfcDatumPlaneOffsetConstraint.OffsetRef
• IpfcDatumPlaneOffsetConstraint.OffsetValue
• IpfcDatumPlaneOffsetCoordSysConstraint.CsysAxis
• IpfcDatumPlaneAngleConstraint.AngleRef
• IpfcDatumPlaneAngleConstraint.AngleValue
• IpfcDatumPlaneSectionConstraint.SectionRef
• IpfcDatumPlaneSectionConstraint.SectionIndex
Thepropertiesofthe IpfcDatumPlaneFeatobjectaredescribed asfollows:
• Flip—Specifies whetherthe datumplanewas flippedduringcreation.Use
| thepropertyIpfcDatumPlaneFeat.Flipto |                        |                |                         |            | determineif            | thedatum |
| ------------------------------------ | ---------------------- | -------------- | ----------------------- | ---------- | ---------------------- | -------- |
| planewas                             | flippedduringcreation. |                |                         |            |                        |          |
| • Constraints—Specifiesa             |                        |                | collectionofconstraints |            | givenbythe             |          |
| IpfcDatumPlaneConstraint             |                        |                |                         | object.The | property               |          |
| IpfcDatumPlaneFeat.Constraints       |                        |                |                         |            | obtainsthecollectionof |          |
| constraintsdefinedfor                |                        | thedatumplane. |                         |            |                        |          |
Usethe propertyIpfcDatumPlaneConstraint.ConstraintType to
| obtainthe | typeof | constraint.Thetypeof |     | constraintisgivenby |     | the |
| --------- | ------ | -------------------- | --- | ------------------- | --- | --- |
IpfcDatumPlaneConstraintTypeenumerated type.The availabletypes
areasfollows:
| • EpfcDTMPLN_THRU—Specifies |     |     |     | theThroughconstraint.The |     |     |
| --------------------------- | --- | --- | --- | ------------------------ | --- | --- |
IpfcDatumPlaneThroughConstraintobjectspecifies thisconstraint.
| Usethe                                     | property |     |     |                    |     |           |
| ------------------------------------------ | -------- | --- | --- | ------------------ | --- | --------- |
| IpfcDatumPlaneThroughConstraint.ThroughRef |          |     |     |                    |     | to getthe |
| referenceselectionhandleforthe             |          |     |     | Throughconstraint. |     |           |
DatumFeatures 211

| • EpfcDTMPLN_NORM—Specifies | theNormalconstraint.The |     |
| --------------------------- | ----------------------- | --- |
IpfcDatumPlaneNormalConstraint objectspecifiesthis constraint.
Usethe propertyIpfcDatumPlaneNormalConstraint.NormalRef
toget thereferenceselectionhandlefor theNormalconstraint.
• EpfcDTMPLN_PRL—SpecifiestheParallelconstraint.
| IpfcDatumPlaneParallelConstraint |     | objectspecifiesthis |
| -------------------------------- | --- | ------------------- |
| constraint.Usethe property       |     |                     |
IpfcDatumPlaneParallelConstraint.ParallelRefto get the
| referenceselectionhandleforthe | Parallelconstraint.      |     |
| ------------------------------ | ------------------------ | --- |
| • EpfcDTMPLN_TANG—Specifies    | theTangentconstraint.The |     |
IpfcDatumPlaneTangentConstraintobjectspecifies thisconstraint.
Usethe property
IpfcDatumPlaneTangentConstraint.TangentRef to getthe
| referenceselectionhandleforthe | Tangentconstraint.      |     |
| ------------------------------ | ----------------------- | --- |
| • EpfcDTMPLN_OFFS—Specifies    | theOffsetconstraint.The |     |
IpfcDatumPlaneOffsetConstraint objectspecifiesthis constraint.
Usethe propertyIpfcDatumPlaneOffsetConstraint.OffsetRef
toget thereferenceselectionhandlefor theOffsetconstraint.Use theproperty
IpfcDatumPlaneOffsetConstraint.OffsetValue to getthe
offsetvalue.
AnOffset constraintwheretheoffsetreferenceisacoordinatesystemis given
bythe IpfcDatumPlaneOffsetCoordSysConstraint object.Use
theproperty
IpfcDatumPlaneOffsetCoordSysConstraint.CsysAxis to get
thereferencecoordinateaxis.
• EpfcDTMPLN_ANG—SpecifiestheAngleconstraint.The
IpfcDatumPlaneAngleConstraintobjectspecifies
thisconstraint.
Usethe propertyIpfcDatumPlaneAngleConstraint.AngleRef to
getthereferenceselectionhandlefor theAngleconstraint.Use theproperty
IpfcDatumPlaneAngleConstraint.AngleValue to getthe angle
value.
| • EpfcDTMPLN_SEC—SpecifiestheSection | constraint.The |     |
| ------------------------------------ | -------------- | --- |
IpfcDatumPlaneSectionConstraintobjectspecifies
thisconstraint.
Usethe property
IpfcDatumPlaneSectionConstraint.SectionRef to getthe
| referenceselectionforthe | Sectionconstraint.Use | theproperty |
| ------------------------ | --------------------- | ----------- |
IpfcDatumPlaneSectionConstraint.SectionIndexto get the
sectionindex.
| 212 |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | ------------------------------- |

| Datum Axis         | Features            |     |              |     |     |
| ------------------ | ------------------- | --- | ------------ | --- | --- |
| Thepropertiesofthe | DatumAxisfeatureare |     | definedinthe |     |     |
IpfcDatumAxisFeatdataobject.
MethodsandPropertiesIntroduced:
• IpfcDatumAxisFeat.Constraints
• IpfcDatumAxisConstraint.ConstraintType
• IpfcDatumAxisConstraint.ConstraintRef
• IpfcDatumAxisFeat.DimConstraints
• IpfcDatumAxisDimensionConstraint.DimOffset
• IpfcDatumAxisDimensionConstraint.DimRef
Thepropertiesofthe IpfcDatumAxisFeat objectaredescribedas follows:
| • Constraints—Specifiesa |     | collectionofconstraints |     | givenbythe |     |
| ------------------------ | --- | ----------------------- | --- | ---------- | --- |
IpfcDatumAxisConstraintobject.Theproperty
| IpfcDatumAxisFeat.Constraintsobtainsthecollection |                         |     |                            |                         | of         |
| ------------------------------------------------- | ----------------------- | --- | -------------------------- | ----------------------- | ---------- |
| constraintsappliedto                              | the DatumAxisfeature.   |     |                            |                         |            |
| Thisobjectcontains                                | thefollowingattributes: |     |                            |                         |            |
| ○ ConstraintType—Specifiesthe                     |                         |     | typeof constraintin        | termsofthe              |            |
| EpfcDatumAxisConstraintTypeenumerated             |                         |     |                            | type.The                | constraint |
| typedeterminesthetypeofdatumaxis.The              |                         |     | constrainttypesare:        |                         |            |
| ◆ EpfcDTMAXIS_NORMAL—Specifies                    |                         |     | theNormaldatumconstraint.  |                         |            |
| ◆ EpfcDTMAXIS_THRU—Specifies                      |                         |     | theThroughdatumconstraint. |                         |            |
| ◆ EpfcDTMAXIS_TANGENT—Specifies                   |                         |     | the                        | Tangentdatumconstraint. |            |
| ◆ EpfcDTMAXIS_CENTER—Specifies                    |                         |     | theCenterdatumconstraint.  |                         |            |
Use thepropertyIpfcDatumAxisConstraint.ConstraintType
| to gettheconstraint | type. |     |     |     |     |
| ------------------- | ----- | --- | --- | --- | --- |
○ ConstraintRef—Specifies thereferenceselectionfor theconstraint.
Use thepropertyIpfcDatumAxisConstraint.ConstraintRef
to getthereferenceselectionhandle.
• DimConstraints—Specifiesacollectionof dimensionconstraints given
| bythe IpfcDatumAxisDimensionConstraint |     |             |                        | object.Theproperty |     |
| -------------------------------------- | --- | ----------- | ---------------------- | ------------------ | --- |
| IpfcDatumAxisFeat.DimConstraints       |     |             | obtainsthecollectionof |                    |     |
| dimensionconstraintsapplied            |     | to theDatum | Axisfeature.           |                    |     |
| ThisIpfcDatumAxisDimensionConstraint   |     |             |                        | objectcontains     | the |
followingattributes:
| DatumFeatures |     |     |     |     | 213 |
| ------------- | --- | --- | --- | --- | --- |

○ DimOffset—Specifies theoffset valueforthe dimensionconstraint.Use
theproperty
| IpfcDatumAxisDimensionConstraint.DimOffset |        |     |     |     | toget the |
| ------------------------------------------ | ------ | --- | --- | --- | --------- |
| offset                                     | value. |     |     |     |           |
○ DimRef—Specifiesthereferenceselectionfor thedimension constraint.
| Use theproperty                         |     |     |     |     |           |
| --------------------------------------- | --- | --- | --- | --- | --------- |
| IpfcDatumAxisDimensionConstraint.DimRef |     |     |     |     | to getthe |
referenceselectionhandle.
| General            | Datum                       | Point | Features |              |     |
| ------------------ | --------------------------- | ----- | -------- | ------------ | --- |
| Thepropertiesofthe | GeneralDatumPointfeatureare |       |          | definedinthe |     |
IpfcDatumPointFeatdataobject.
MethodsandPropertiesIntroduced:
• IpfcDatumPointFeat.FeatName
• IpfcDatumPointFeat.GetPoints()
• IpfcGeneralDatumPoint.Name
• IpfcGeneralDatumPoint.PlaceConstraints
• IpfcGeneralDatumPoint.DimConstraints
• IpfcDatumPointConstraint.ConstraintRef
• IpfcDatumPointConstraint.ConstraintType
• IpfcDatumPointConstraint.Value
Thepropertiesofthe IpfcDatumPointFeatobjectaredescribed asfollows:
• FeatName—Specifiesthenameof theGeneralDatum Pointfeature.Usethe
| propertyIpfcDatumPointFeat.FeatNameto |     |     |     | getthe | name. |
| ------------------------------------- | --- | --- | --- | ------ | ----- |
• GeneralDatumPoints—Specifiesacollection ofgeneraldatumpoints
| givenby                        | theIpfcGeneralDatumPoint |     | object.Usethe |           | method       |
| ------------------------------ | ------------------------ | --- | ------------- | --------- | ------------ |
| IpfcDatumPointFeat.GetPoints() |                          |     | to            | obtainthe | collectionof |
generaldatumpoints.TheIpfcGeneralDatumPoint objectconsists of
thefollowingattributes:
○ Name—Specifies thenameof thegeneraldatumpoint.Use theproperty
| IpfcGeneralDatumPoint.Name |     |     | to getthename. |     |     |
| -------------------------- | --- | --- | -------------- | --- | --- |
○ PlaceConstraints—Specifiesa collectionofplacementconstraints
| givenby | theIpfcDatumPointPlacementConstraint |     |     |     | object.Use |
| ------- | ------------------------------------ | --- | --- | --- | ---------- |
thepropertyIpfcGeneralDatumPoint.PlaceConstraints to
| obtainthecollectionof |     | placementconstraints. |     |                                 |     |
| --------------------- | --- | --------------------- | --- | ------------------------------- | --- |
| 214                   |     |                       |     | Creo®ParametricVBAPIUser’sGuide |     |

| ○ DimConstraints—Specifiesa |     |                                      |     | collectionofdimensionconstraints |     |            |
| --------------------------- | --- | ------------------------------------ | --- | -------------------------------- | --- | ---------- |
| givenby                     |     | theIpfcDatumPointDimensionConstraint |     |                                  |     | object.Use |
thepropertyIpfcGeneralDatumPoint.DimConstraints to
| obtainthecollectionof |     | dimensionconstraints. |     |       |     |     |
| --------------------- | --- | --------------------- | --- | ----- | --- | --- |
| Theconstraintsfor     |     | adatumpointaregiven   |     | bythe |     |     |
IpfcDatumPointConstraintobject.This objectcontainsthefollowing
attributes:
• ConstraintRef—Specifiesthe referenceselectionforthe datumpoint
| constraint.Usethe                      |     | property |     |     |     |                     |
| -------------------------------------- | --- | -------- | --- | --- | --- | ------------------- |
| IpfcDatumPointConstraint.ConstraintRef |     |          |     |     |     | to getthe reference |
selectionhandle.
• ConstraintType—Specifiesthetypeof datumpoint constraint.interms
| oftheIpfcDatumPointConstraintType               |     |     |     |     | enumerated | type.Use the |
| ----------------------------------------------- | --- | --- | --- | --- | ---------- | ------------ |
| propertyIpfcDatumPointConstraint.ConstraintType |     |     |     |     |            | to getthe    |
constrainttype.
• Value—Specifiestheconstraint referencevaluewith respectto thedatum
point.Usethe propertyIpfcDatumPointConstraint.Value to getthe
| valueofthe                                          |            | constraintreferencewith |                      | respectto    | thedatumpoint. |     |
| --------------------------------------------------- | ---------- | ----------------------- | -------------------- | ------------ | -------------- | --- |
| TheIpfcDatumPointPlacementConstraint                |            |                         |                      |              | and            |     |
| IpfcDatumPointDimensionConstraintobjectsinheritfrom |            |                         |                      |              |                | the |
| IpfcDatumPointConstraintobject.Use                  |            |                         |                      | themethodsof |                | the |
| IpfcDatumPointFeatobjectfor                         |            |                         | theinheritedobjects. |              |                |     |
| Datum                                               | Coordinate |                         | System               |              | Features       |     |
Thepropertiesofthe DatumCoordinateSystem featurearedefinedin the
IpfcCoordSysFeatobject.
MethodsandPropertiesIntroduced:
• IpfcCoordSysFeat.OriginConstraints
• IpfcDatumCsysOriginConstraint.OriginRef
• IpfcCoordSysFeat.DimensionConstraints
• IpfcDatumCsysDimensionConstraint.DimRef
• IpfcDatumCsysDimensionConstraint.DimValue
• IpfcDatumCsysDimensionConstraint.DimConstraintType
• IpfcCoordSysFeat.OrientationConstraints
• IpfcDatumCsysOrientMoveConstraint.OrientMoveConstraintType
• IpfcDatumCsysOrientMoveConstraint.OrientMoveValue
• IpfcCoordSysFeat.IsNormalToScreen
DatumFeatures 215

• IpfcCoordSysFeat.OffsetType
• IpfcCoordSysFeat.OnSurfaceType
• IpfcCoordSysFeat.OrientByMethod
Thepropertiesofthe IpfcCoordSysFeatobjectaredescribed asfollows:
• OriginConstraints—Specifies acollectionof originconstraintsgiven
| bythe IpfcDatumCsysOriginConstraintobject.Use |     | theproperty              |
| --------------------------------------------- | --- | ------------------------ |
| IpfcCoordSysFeat.OriginConstraints            |     | to obtainthecollectionof |
originconstraints forthecoordinatesystem.This objectcontainsthe following
attribute:
○ OriginRef—Specifies theselectionreferenceforthe origin.Use the
methodproperty IpfcDatumCsysOriginConstraint.OriginRef
to gettheselection referencehandle.
| • DimensionConstraints—Specifiesa | collectionofdimension |     |
| --------------------------------- | --------------------- | --- |
constraintsgivenby theIpfcDatumCsysDimensionConstraint
object.Use theproperty
| IpfcCoordSysFeat.DimensionConstraints |     | toobtainthe |
| ------------------------------------- | --- | ----------- |
collectionofdimension constraintsfor thecoordinatesystem.Thisobject
containsthe followingattributes:
○ DimRef—Specifiesthereferenceselectionfor thedimension constraint.
Use theproperty
| IpfcDatumCsysDimensionConstraint.DimRef |     | to getthe |
| --------------------------------------- | --- | --------- |
referenceselectionhandle.
| ○ DimValue—Specifiesthevalue | ofthereference. | Use theproperty |
| ---------------------------- | --------------- | --------------- |
IpfcDatumCsysDimensionConstraint.DimValue to getthe
value.
○ DimConstraintType—Specifies thetypeofdimension constraintin
| termsof theEpfcDatumCsysDimConstraintType |     | enumerated |
| ----------------------------------------- | --- | ---------- |
type.Use theproperty
IpfcDatumCsysDimensionConstraint.DimConstraintType
| to gettheconstraint type.The | constrainttypesare: |     |
| ---------------------------- | ------------------- | --- |
EpfcDTMCSYS_DIM_OFFSET—Specifies
| ◆                                   |                         | theoffsettypeconstraint. |
| ----------------------------------- | ----------------------- | ------------------------ |
| ◆ EpfcDTMCSYS_DIM_ALIGN—Specifies   |                         | the aligntypeconstraint. |
| • OrientationConstraints—Specifiesa | collectionoforientation |                          |
constraintsgivenby theIpfcDatumCsysOrientMoveConstraint
object.Use theproperty
| IpfcCoordSysFeat.OrientationConstraints |                                 | toobtainthe |
| --------------------------------------- | ------------------------------- | ----------- |
| 216                                     | Creo®ParametricVBAPIUser’sGuide |             |

collectionoforientation constraintsforthe coordinatesystem.Thisobject
containsthe followingattributes:
| ○ OrientMoveConstraintType—Specifies  |                          |     |     | thetypeoforientationfor |                 |     |     |
| ------------------------------------- | ------------------------ | --- | --- | ----------------------- | --------------- | --- | --- |
| theconstraint.The                     | orientationtypeisgivenby |     |     | the                     |                 |     |     |
| EpfcDatumCsysOrientMoveConstraintType |                          |     |     |                         | enumeratedtype. |     |     |
Use theproperty
IpfcDatumCsysOrientMoveConstraint.OrientMoveCon
| straintType | toget theorientationtype. |     |     |     |     |     |     |
| ----------- | ------------------------- | --- | --- | --- | --- | --- | --- |
○ OrientMoveValue—Specifies thereferencevaluefor theconstraint.
Use theproperty
IpfcDatumCsysOrientMoveConstraint.OrientMoveValue
to getthereferencevalue.
• IsNormalToScreen—Specifiesif thecoordinatesystemisnormal to
| screen.Usethe propertyIpfcCoordSysFeat.IsNormalToScreento |     |            |     |         |     |     |     |
| --------------------------------------------------------- | --- | ---------- | --- | ------- | --- | --- | --- |
| determineifthecoordinatesystem                            |     | isnormalto |     | screen. |     |     |     |
• OffsetType—Specifiestheoffsettypeofthecoordinatesystemin termsof
| theEpfcDatumCsysOffsetType |     | enumerated |     | type.Use | theproperty |     |     |
| -------------------------- | --- | ---------- | --- | -------- | ----------- | --- | --- |
IpfcCoordSysFeat.OffsetTypeto getthe offsettype. Theoffsettypes
are:
| ○ EpfcDTMCSYS_OFFSET_CARTESIAN—Specifies       |     |                          |                         |     | acartesian |     |     |
| ---------------------------------------------- | --- | ------------------------ | ----------------------- | --- | ---------- | --- | --- |
| coordinatesystemthathasbeendefinedbysettingthe |     |                          |                         |     | valuesfor  | the |     |
| EpfcDTMCSYS_MOVE_TRAN_X,                       |     | EpfcDTMCSYS_MOVE_TRAN_Y, |                         |     |            |     | and |
| EpfcDTMCSYS_MOVE_TRAN_Zor                      |     |                          | EpfcDTMCSYS_MOVE_ROT_X, |     |            |     |     |
| EpfcDTMCSYS_MOVE_ROT_Y,and                     |     |                          | EpfcDTMCSYS_MOVE_ROT_Z  |     |            |     |     |
orientationconstants.
| ○ EpfcDTMCSYS_OFFSET_CYLINDRICAL—Specifies     |     |     |                       |            | acylindrical |     |     |
| ---------------------------------------------- | --- | --- | --------------------- | ---------- | ------------ | --- | --- |
| coordinatesystemthathasbeendefinedbysettingthe |     |     |                       |            | valuesfor    | the |     |
| EpfcDTMCSYS_MOVE_RAD,EpfcDTMCSYS_MOVE_THETA,   |     |     |                       |            |              | and |     |
| EpfcDTMCSYS_MOVE_TRAN_ZI                       |     |     | orientationconstants. |            |              |     |     |
| ○ EpfcDTMCSYS_OFFSET_SPHERICAL—Specifies       |     |     |                       |            | aspherical   |     |     |
| coordinatesystemthathasbeendefinedbysettingthe |     |     |                       |            | valuesfor    | the |     |
| EpfcDTMCSYS_MOVE_RAD,EpfcDTMCSYS_MOVE_THETA,   |     |     |                       |            |              | and |     |
| EpfcDTMCSYS_MOVE_TRAN_PHIorientation           |     |     |                       | constants. |              |     |     |
• OnSurfaceType—Specifiesthe onsurface typeforthe coordinatesystem
| interms oftheEpfcDatumCsysOffsetType   |     |                           |           | enumeratedtype.Use |                  |         | the |
| -------------------------------------- | --- | ------------------------- | --------- | ------------------ | ---------------- | ------- | --- |
| propertyIpfcCoordSysFeat.OnSurfaceType |     |                           |           |                    | to gettheon      | surface |     |
| typepropertyofthecoordinatesystem.The  |     |                           | onsurface |                    | typesare:        |         |     |
| ○ EpfcDTMCSYS_ONSURF_LINEAR—Specifies  |     |                           |           | a                  | coordinatesystem |         |     |
| placedon theselectedsurfaceby          |     | usingtwolineardimensions. |           |                    |                  |         |     |
| DatumFeatures                          |     |                           |           |                    |                  |         | 217 |

| ○ EpfcDTMCSYS_ONSURF_RADIAL—Specifies |     |     |     | a   | coordinatesystem |     |
| ------------------------------------- | --- | --- | --- | --- | ---------------- | --- |
placedon theselectedsurfaceby usinga lineardimension andan angular
| dimension.The                      | radiusvalueisused |     | tospecify | thelineardimension.        |        |     |
| ---------------------------------- | ----------------- | --- | --------- | -------------------------- | ------ | --- |
| ○ EpfcDTMCSYS_ONSURF_DIAMETER—This |                   |     |           | typeissimilar              | to the |     |
| EpfcDTMCSYS_ONSURF_RADIALtype,     |                   |     |           | exceptthatthediametervalue |        |     |
isused tospecify thelineardimension.It isavailableonlywhen planar
| surfacesare                                | usedasthe | reference. |     |          |                    |     |
| ------------------------------------------ | --------- | ---------- | --- | -------- | ------------------ | --- |
| • OrientByMethod—Specifiestheorientation   |           |            |     | methodin | termsofthe         |     |
| EpfcDatumCsysOrientByMethodenumeratedtype. |           |            |     |          | Use theproperty    |     |
| IpfcCoordSysFeat.OrientByMethodto          |           |            |     | getthe   | orientationmethod. |     |
Theavailableorientationtypesare:
| ○ EpfcDTMCSYS_ORIENT_BY_SEL_REFS—Specifies |     |     |     |     | theorientation | by  |
| ------------------------------------------ | --- | --- | --- | --- | -------------- | --- |
selectedreferences.
| ○ EpfcDTMCSYS_ORIENT_BY_SEL_CSYS_AXES—Specifies |                  |     |       |                                 | the |     |
| ----------------------------------------------- | ---------------- | --- | ----- | ------------------------------- | --- | --- |
| orientationby                                   | coordinatesystem |     | axes. |                                 |     |     |
| 218                                             |                  |     |       | Creo®ParametricVBAPIUser’sGuide |     |     |

17
Geometry Evaluation
GeometryTraversal.................................................................................................220
CurvesandEdges...................................................................................................221
Contours.................................................................................................................224
Surfaces.................................................................................................................225
Axes,CoordinateSystems,andPoints......................................................................229
Interference............................................................................................................229
Thischapterdescribes geometryrepresentationanddiscusses howtoevaluate
geometryusingthe VBAPI.
219

| Geometry                       |     | Traversal |     |            |            |     |     |     |
| ------------------------------ | --- | --------- | --- | ---------- | ---------- | --- | --- | --- |
| • Asimplerectangularfacehasone |     |           |     | contourand | fouredges. |     |     |     |
• Acontourwilltraverse aboundarysothatthepart faceisalwayson theright-
| handside      | (RHS).For | anexternalcontourthedirection |     |                      |     | oftraversal |     | is  |
| ------------- | --------- | ----------------------------- | --- | -------------------- | --- | ----------- | --- | --- |
| clockwise.For | an        | internalcontourthe            |     | directionoftraversal |     |             | is  |     |
counterclockwise.
• Ifapart isextruded froma sketchthathasaU-shaped crosssectiontherewill
| beseparate | surfacesateachleg |     | oftheU-channel. |     |     |     |     |     |
| ---------- | ----------------- | --- | --------------- | --- | --- | --- | --- | --- |
• Ifapart isextruded froma sketchthathasasquare-shapedcrosssection,anda
slotfeatureisthencut intothe partto makeitlook likeaU-channel, therewill
beone surfaceacross thelegs oftheU-channel. Theoriginal surface ofthe
| partis representedas       |       | onesurfacewith |                     | acutthrough |     | it. |     |     |
| -------------------------- | ----- | -------------- | ------------------- | ----------- | --- | --- | --- | --- |
| Geometry                   | Terms |                |                     |             |     |     |     |     |
| Followingaredefinitionsfor |       |                | somegeometricterms: |             |     |     |     |     |
• Surface—Anideal geometricrepresentation,thatis,an infiniteplane.
• Face—Atrimmedsurface.Afacehasoneormorecontours.
• Contour—Aclosedloop on aface.Acontourconsists ofmultiple edges.A
| contourcanbelong             |     | toone | faceonly. |          |     |     |     |     |
| ---------------------------- | --- | ----- | --------- | -------- | --- | --- | --- | --- |
| • Edge—Theboundaryofatrimmed |     |       |           | surface. |     |     |     |     |
Anedgeofasolidis theintersectionof twosurfaces.The edgebelongs to those
twosurfacesandto twocontours. Anedgeofa datumsurface canbe eitherthe
intersectionoftwodatumsurfaces ortheexternalboundaryof thesurface.
Iftheedgeisthe intersectionoftwodatumsurfaces itwillbelong to thosetwo
surfacesandto twocontours. Iftheedgeisthe externalboundary ofthedatum
| surfaceitwillbelongto |     | thatsurfacealone |     | andto      | asinglecontour. |       |     |     |
| --------------------- | --- | ---------------- | --- | ---------- | --------------- | ----- | --- | --- |
| Traversing            | the | Geometry         |     | of a Solid |                 | Block |     |     |
MethodsIntroduced:
• IpfcModelItemOwner.ListItems()
• IpfcSurface.ListContours()
• IpfcContour.ListElements()
| Totraversethe | geometry,follow |     | thesesteps: |     |     |                                 |     |     |
| ------------- | --------------- | --- | ----------- | --- | --- | ------------------------------- | --- | --- |
| 220           |                 |     |             |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |

| 1. Startingatthe               | top-levelmodel,use |     |                   |     |     |
| ------------------------------ | ------------------ | --- | ----------------- | --- | --- |
| IpfcModelItemOwner.ListItems() |                    |     | withan argumentof |     |     |
ModelItemType.ITEM_SURFACE.
2. UseIpfcSurface.ListContours() to list thecontourscontainedin a
specifiedsurface.
3. UseIpfcContour.ListElements() to list theedgescontainedin the
contour.
| Curves | and Edges |     |     |     |     |
| ------ | --------- | --- | --- | --- | --- |
Datumcurves,surfaceedges, andsolidedgesarerepresented inthe samewayin
theVBAPI. You canget edgesthrough geometrytraversal orgeta listof edges
| usingthe methodspresented |     | inthe chapterModelItemonpage188. |     |     |     |
| ------------------------- | --- | -------------------------------- | --- | --- | --- |
| The t Parameter           |     |                                  |     |     |     |
Thegeometryof eachedgeor curveisrepresented asasetof threeparametric
equationsthatrepresent thevalues ofx,y, andzasfunctions ofan independent
| parameter,t.The | tparameter |               |                 |            |       |
| --------------- | ---------- | ------------- | --------------- | ---------- | ----- |
|                 |            | variesfrom0.0 | atthestartofthe | curveto1.0 | atthe |
endofit.
Thefollowingfigureillustratescurveandedgeparameterization.
| Curve and          | Edge Types     |                            |     |     |     |
| ------------------ | -------------- | -------------------------- | --- | --- | --- |
| Solidedgesand      | datumcurvescan | be anyofthefollowingtypes: |     |     |     |
| GeometryEvaluation |                |                            |     |     | 221 |

| • LINE—Astraightlinerepresented |                    |                         | bytheinterface |                      | IpfcLine. |           |
| ------------------------------- | ------------------ | ----------------------- | -------------- | -------------------- | --------- | --------- |
| • ARC—Acircular                 | curverepresentedby |                         |                | theinterfaceIpfcArc. |           |           |
| • SPLINE—Anonuniform            |                    | cubicspline,represented |                |                      | bythe     | interface |
IpfcSpline.
• B-SPLINE—Anonuniform rationalB-spline curveoredge, representedby the
interfaceIpfcBSpline.
| • COMPOSITE                            | CURVE—Acombinationoftwoor |     |     |     | morecurves,represented |             |
| -------------------------------------- | ------------------------- | --- | --- | --- | ---------------------- | ----------- |
| bythe interfaceIpfcCompositeCurve.This |                           |     |     |     | isusedfor              | datumcurves |
only.
Seetheappendix GeometryRepresentationson page374forthe parameterization
| ofeachcurvetype.Todeterminewhat  |     |     | typeof                  | curveaIpfcEdgeor |     |     |
| -------------------------------- | --- | --- | ----------------------- | ---------------- | --- | --- |
| IpfcCurveobjectrepresents,usethe |     |     | Javainstanceofoperator. |                  |     |     |
Becauseeachcurveclassinherits fromIpfcGeomCurve, youcanuseallthe
| evaluationmethodsin | IpfcGeomCurveon |     |     | anyedgeorcurve. |     |     |
| ------------------- | --------------- | --- | --- | --------------- | --- | --- |
Thefollowingcurvetypesarenotusedin solidgeometryandarereservedfor
futureexpansion:
• CIRCLE(IpfcCircle)
• ELLIPSE(IpfcEllipse)
| • POLYGON | (IpfcPolygon) |     |     |     |     |     |
| --------- | ------------- | --- | --- | --- | --- | --- |
• ARROW(IpfcArrow)
• TEXT(IpfcText)
| Evaluation | of Curves | and | Edges |     |     |     |
| ---------- | --------- | --- | ----- | --- | --- | --- |
MethodsIntroduced:
• IpfcGeomCurve.Eval3DData()
• IpfcGeomCurve.EvalFromLength()
• IpfcGeomCurve.EvalParameter()
• IpfcGeomCurve.EvalLength()
• IpfcGeomCurve.EvalLengthBetween()
Themethodsin IpfcGeomCurve provideinformationaboutany curveoredge.
| ThemethodIpfcGeomCurve.Eval3DData() |     |     |     |     | returnsa |     |
| ----------------------------------- | --- | --- | --- | --- | -------- | --- |
IpfcCurveXYZDataobjectwith informationonthe pointrepresented bythe
| inputparametert.The | methodIpfcGeomCurve.EvalFromLength() |     |     |     |     |     |
| ------------------- | ------------------------------------ | --- | --- | --- | --- | --- |
returnsasimilar objectwith informationon thepoint thatis aspecifieddistance
| fromthestarting | point. |     |     |     |                                 |     |
| --------------- | ------ | --- | --- | --- | ------------------------------- | --- |
| 222             |        |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |

| ThemethodIpfcGeomCurve.EvalParameter() |                    |     | returnsthet | parameter |
| -------------------------------------- | ------------------ | --- | ----------- | --------- |
| thatrepresents theinput                | IpfcPoint3Dobject. |     |             |           |
BothIpfcGeomCurve.EvalLength() and
IpfcGeomCurve.EvalLengthBetween()returnnumericalvalues forthe
lengthof thecurveoredge.
| Solid Edge | Geometry |     |     |     |
| ---------- | -------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcEdge.Surface1
• IpfcEdge.Surface2
• IpfcEdge.Edge1
• IpfcEdge.Edge2
• IpfcEdge.EvalUV()
• IpfcEdge.GetDirection()
Note
| The methodsin                  | theinterfaceIpfcEdgeprovideinformation |                             |           | onlyfor |
| ------------------------------ | -------------------------------------- | --------------------------- | --------- | ------- |
| solidorsurface                 | edges.                                 |                             |           |         |
| ThepropertiesIpfcEdge.Surface1 |                                        | andEdge.Surface2            | returnthe |         |
| surfacesboundedby              | thisedge.The                           | propertiesIpfcEdge.Edge1and |           |         |
IpfcEdge.Edge2returnthenextedgesin thetwocontoursthatcontainthis
edge.
ThemethodIpfcEdge.EvalUV() evaluates geometryinformationbasedon
| theUVparametersofoneof | theboundingsurfaces. |     |     |     |
| ---------------------- | -------------------- | --- | --- | --- |
ThemethodIpfcEdge.GetDirection() returnsapositive 1ifthe edgeis
parameterizedin thesamedirection asthecontainingcontour,and–1 if theedge
| isparameterizedoppositeto | thecontainingcontour. |     |     |     |
| ------------------------- | --------------------- | --- | --- | --- |
Curve Descriptors
Acurvedescriptoris adataobjectthatdescribesthe geometryofacurveor edge.
Acurvedescriptordescribes thegeometryof acurvewithoutbeingapart ofa
specificmodel.
MethodsIntroduced:
• IpfcGeomCurve.GetCurveDescriptor()
GeometryEvaluation 223

• IpfcGeomCurve.GetNURBSRepresentation()
Note
|     | Togetgeometricinformationfor |           | an edge,access |     | the |
| --- | ---------------------------- | --------- | -------------- | --- | --- |
|     | IpfcCurveDescriptor          | objectfor | oneedgeusing   |     |     |
IpfcGeomCurve.GetCurveDescriptor().
| ThemethodIpfcGeomCurve.GetCurveDescriptor() |     |     |     |     | returnsacurve’s |
| ------------------------------------------- | --- | --- | --- | --- | --------------- |
geometryasadataobject.
ThemethodIpfcGeomCurve.GetNURBSRepresentation() returnsa
| Non-UniformRationalB-Spline |     | Representationof |     | acurve. |     |
| --------------------------- | --- | ---------------- | --- | ------- | --- |
Contours
MethodsandPropertiesIntroduced:
• IpfcSurface.ListContours()
• IpfcContour.InternalTraversal
• IpfcContour.FindContainingContour()
• IpfcContour.EvalArea()
• IpfcContour.EvalOutline()
• IpfcContour.VerifyUV()
Contoursareaseriesofedgesthatcompletelybounda surface.Acontourisnota
IpfcModelItem.Youcannotgetcontoursusing themethods thatgetdifferent
| typesof                      | IpfcModelItem.Use | themethod |                 |     |                 |
| ---------------------------- | ----------------- | --------- | --------------- | --- | --------------- |
| IpfcSurface.ListContours()to |                   |           | getcontoursfrom |     | theircontaining |
surfaces.
| ThepropertyIpfcContour.InternalTraversal |     |     |     |     | returnsa |
| ---------------------------------------- | --- | --- | --- | --- | -------- |
EpfcContourTraversalenumeratedtypethatidentifies whethera given
| contouris | onthe outsideorinsideof |     | acontainingsurface. |     |     |
| --------- | ----------------------- | --- | ------------------- | --- | --- |
Usethe methodIpfcContour.FindContainingContour() tofind the
| contourthatentirelyenclosesthe |     | specifiedcontour. |     |     |     |
| ------------------------------ | --- | ----------------- | --- | --- | --- |
ThemethodIpfcContour.EvalArea() providestheareaenclosed bythe
contour.
ThemethodIpfcContour.EvalOutline() returnsthepoints thatmake up
theboundingrectangle ofthecontour.
| 224 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

Usethe methodIpfcContour.VerifyUV()to determinewhetherthegiven
IpfcUVParamsargument liesinside thecontour,on theboundary, oroutsidethe
contour.
Surfaces
| UsingtheVB | API youaccessdatumandsolidsurfacesin |     |     |     | thesameway. |
| ---------- | ------------------------------------ | --- | --- | --- | ----------- |
UV Parameterization
Asurfacein CreoParametricis describedas aseriesofparametric equations
wheretwoparameters, uandv,determinethex,y, andzcoordinates. Unlike the
edgeparameter,t,theseparametersneednot startat0.0,noraretheylimitedto
1.0.
| Thefigureon     | thefollowingpageillustratessurface |                    |          | parameterization.   |     |
| --------------- | ---------------------------------- | ------------------ | -------- | ------------------- | --- |
| Surface         | Types                              |                    |          |                     |     |
| Surfaceswithin  | Creo                               | Parametriccanbe    | anyofthe | followingtypes:     |     |
| • PLANE—Aplanar |                                    | surfacerepresented | bythe    | interfaceIpfcPlane. |     |
GeometryEvaluation 225

| • CYLINDER—Acylindricalsurface |     |     | representedby | the |     |
| ------------------------------ | --- | --- | ------------- | --- | --- |
interfaceIpfcCylinder.
• CONE—Aconicsurface regionrepresented bytheinterfaceIpfcCone.
• TORUS—A toroidalsurface regionrepresented bythe interfaceIpfcTorus.
• REVOLVEDSURFACE—Generatedby revolvinga curveaboutan axis.This
| isrepresented | bythe | interfaceIpfcRevSurface. |     |     |     |
| ------------- | ----- | ------------------------ | --- | --- | --- |
• RULEDSURFACE—Generatedby interpolatinglinearlybetweentwocurve
| entities.This | isrepresented | bythe | interfaceIpfcRuledSurface. |     |     |
| ------------- | ------------- | ----- | -------------------------- | --- | --- |
• TABULATEDCYLINDER—Generatedby extrudingacurvelinearly.This is
theinterfaceIpfcTabulatedCylinder.
representedby
• COONSPATCH—A coons patchisused toblend surfacestogether. Itis
| representedby | theinterfaceIpfcCoonsPatch. |     |     |     |     |
| ------------- | --------------------------- | --- | --- | --- | --- |
• FILLETSURFACE—A filletedsurfaceis foundwherea roundor filletis
placedona curvededgeoran edgewitha non-consistantarcradii.On a
straightedgeacylinder isusedto represent afillet.This isrepresented by the
| classinterface   | IpfcFilletedSurface. |                          |     |                   |     |
| ---------------- | -------------------- | ------------------------ | --- | ----------------- | --- |
| • SPLINESURFACE— |                      | Anonuniformbicubicspline |     | surfacethatpasses |     |
througha gridwith tangentvectors givenateachpoint.Thisis representedby
theinterfaceIpfcSplineSurface.
• NURBS SURFACE—ANURBS surface isdefinedby basicfunctions(inu
andv),expandablearraysofknots, weights,andcontrolpoints.This is
| representedby               | theclass | interfaceIpfcNURBSSurface. |                           |     |     |
| --------------------------- | -------- | -------------------------- | ------------------------- | --- | --- |
| • CYLINDRICALSPLINESURFACE— |          |                            | Acylindricalsplinesurface |     | isa |
nonuniformbicubic splinesurfacethatpasses through agridwith tangent
| vectorsgiven | ateachpoint.This |     | isrepresented | bythe interface |     |
| ------------ | ---------------- | --- | ------------- | --------------- | --- |
IpfcCylindricalSplineSurface.
Todeterminewhichtypeofsurface aIpfcSurface objectrepresents,access
| thesurfacetypeusing | IpfcSurface.GetSurfaceType(). |     |     |     |     |
| ------------------- | ----------------------------- | --- | --- | --- | --- |
| Surface Information |                               |     |     |     |     |
MethodsIntroduced:
• IpfcSurface.GetSurfaceType()
• IpfcSurface.GetXYZExtents()
• IpfcSurface.GetUVExtents()
• IpfcSurface.GetOrientation()
| 226 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

ThemethodIpfcSurface.GetSurfaceType() returns thetypeofthe
surfaceusing theenumerateddatatypeEpfcSurfaceType andthevalid values
are:
• EpfcSURFACE_PLANE
• EpfcSURFACE_CYLINDER
• EpfcSURFACE_CONE
• EpfcSURFACE_TORUS
• EpfcSURFACE_RULED
• EpfcSURFACE_REVOLVED
• EpfcSURFACE_TABULATED_CYLINDER
• EpfcSURFACE_FILLET
• EpfcSURFACE_COONS_PATCH
• EpfcSURFACE_SPLINE
• EpfcSURFACE_NURBS
• EpfcSURFACE_CYLINDRICAL_SPLINE
• EpfcSURFACE_FOREIGN
• EpfcSURFACE_SPL2DER
ThemethodIpfcSurface::GetXYZExtents() returns theXYZpointsat
thecornersof thesurface.
ThemethodIpfcSurface::GetUVExtents() returnstheUV parametersat
thecornersof thesurface.
ThemethodIpfcSurface::GetOrientation() returnstheorientation of
thesurfaceusing theenumerated datatypeEpfcSurfaceOrientation and
thevalidvalues are:
• EpfcSURFACEORIENT_NONE—Surfacethatdoes notneedorientation.For
example,a solidsurfaceneedsorientationandtherefore cannotbe specified.
• EpfcSURFACEORIENT_OUTWARD—Surface thathasorientedoutward
awayfrom thesolidmodel. duXdv pointsoutward.
• EpfcSURFACEORIENT_INWARD—Surfacethathasoriented inwardtoward
thesolidmodel.du Xdv pointsinward.
Evaluation of Surfaces
Surfacemethodsallow youtousemultiple surfaceinformation tocalculate,
evaluate,determine, andexaminesurface functionsandproblems.
MethodsandPropertiesIntroduced:
• IpfcSurface.OwnerQuilt
• IpfcSurface.EvalClosestPoint()
GeometryEvaluation 227

• IpfcSurface.EvalClosestPointOnSurface()
• IpfcSurface.Eval3DData()
• IpfcSurface.EvalParameters()
• IpfcSurface.EvalArea()
• IpfcSurface.EvalDiameter()
• IpfcSurface.EvalPrincipalCurv()
• IpfcSurface.VerifyUV()
• IpfcSurface.EvalMaximum()
• IpfcSurface.EvalMinimum()
• IpfcSurface.ListSameSurfaces
| ThepropertyIpfcSurface.OwnerQuilt |     |     |     | Quiltobjectthat |
| --------------------------------- | --- | --- | --- | --------------- |
returnsthe
containsthedatumsurface.
| ThemethodIpfcSurface.EvalClosestPoint()   |                |           |     | projects athree- |
| ----------------------------------------- | -------------- | --------- | --- | ---------------- |
| dimensionalpointonto                      | thesurface.Use | themethod |     |                  |
| IpfcSurface.EvalClosestPointOnSurface()to |                |           |     | determinewhether |
thespecifiedthree-dimensionalpointison thesurface, withinthe accuracyof the
part.Ifitis,themethod returnsthepointthatis exactlyon thesurface.Otherwise
themethodreturnsnull.
| ThemethodIpfcSurface.Eval3DData() |     |     | returnsa | IpfcSurfXYZData |
| --------------------------------- | --- | --- | -------- | --------------- |
objectthatcontains informationaboutthe surfaceatthespecifieduand v
parameters.Themethod IpfcSurface.EvalParameters() returnstheu
andvparametersthatcorrespondto thespecifiedthree-dimensionalpoint.
ThemethodIpfcSurface.EvalArea() returnsthe areaof thesurface,
| whereasIpfcSurface.EvalDiameter() |     |     | returnsthe | diameterofthe |
| --------------------------------- | --- | --- | ---------- | ------------- |
surface.Ifthediameter variestheoptionalIpfcUVParamsargumentidentifies
| wherethe diametershouldbe                | evaluated. |     |     |          |
| ---------------------------------------- | ---------- | --- | --- | -------- |
| ThemethodIpfcSurface.EvalPrincipalCurv() |            |     |     | returnsa |
IpfcCurvatureDataobjectwith informationregarding thecurvatureof the
| surfaceatthespecifieduandv            | parameters. |               |                     |     |
| ------------------------------------- | ----------- | ------------- | ------------------- | --- |
| Usethe methodIpfcSurface.VerifyUV()to |             |               | determinewhetherthe |     |
| IpfcUVParamsareactuallywithin         |             | theboundaryof | thesurface.         |     |
| ThemethodsIpfcSurface.EvalMaximum()   |             |               | and                 |     |
IpfcSurface.EvalMinimum()returnthethree-dimensional pointonthe
surfacethatisthe furthestin the directionof(or awayfrom)the specifiedvector.
ThemethodIpfcSurface.ListSameSurfaces() identifies othersurfaces
| thatare tangentandconnect | tothe | givensurface. |                                 |     |
| ------------------------- | ----- | ------------- | ------------------------------- | --- |
| 228                       |       |               | Creo®ParametricVBAPIUser’sGuide |     |

| Surface | Descriptors |     |     |
| ------- | ----------- | --- | --- |
Asurfacedescriptorisa dataobjectthatdescribes theshapeandgeometryof a
specifiedsurface. Asurfacedescriptorallowsyouto describea surfacein 3D
| withoutanowner | ID. |     |     |
| -------------- | --- | --- | --- |
MethodsIntroduced:
• IpfcSurface.GetSurfaceDescriptor()
• IpfcSurface.GetNURBSRepresentation()
ThemethodIpfcSurface.GetSurfaceDescriptor()
returnsasurfaces
geometryasadataobject.
ThemethodIpfcSurface.GetNURBSRepresentation() returnsaNon-
| UniformRational | B-SplineRepresentation | ofasurface. |            |
| --------------- | ---------------------- | ----------- | ---------- |
| Axes,           | Coordinate             | Systems,    | and Points |
Coordinateaxes,datumpoints,andcoordinatesystems areallmodelitems.Use
themethodsthatreturnIpfcModelItems toget oneofthesegeometryobjects.
Refertothe chapterModelItemonpage188for additionalinformation.
| Evaluation | of ModelItems |     |     |
| ---------- | ------------- | --- | --- |
PropertiesIntroduced:
• IpfcAxis.Surf
• IpfcCoordSystem.CoordSys
• IpfcPoint.Point
ThepropertyIpfcAxis.Surf returnsthe revolvedsurface thatusesthe axis.
| ThepropertyIpfcCoordSystem.CoordSys |     | returnsthe |     |
| ----------------------------------- | --- | ---------- | --- |
IpfcTransform3Dobject(which includestheorigin andx-,y-,andz- axes)
| thatdefines | thecoordinatesystem. |     |     |
| ----------- | -------------------- | --- | --- |
ThepropertyIpfcPoint.Point
|     |     | returnsthe xyzcoordinatesofthedatum |     |
| --- | --- | ----------------------------------- | --- |
point.
Interference
CreoParametricassembliescan containinterferences betweencomponentswhen
constraintbycertain rulesdefinedby theuser.The IpfcInterference
moduleallowstheuserto detectandanalyze anyinterferences withinthe
assembly. Theanalysisof thisfunctionalityshouldbe looked atfromtwo
standpoints:globaland selectionbasedanalysis.
GeometryEvaluation 229

MethodsandPropertiesIntroduced:
• CMpfcInterference.CreateGlobalEvaluator()
• IpfcGlobalEvaluator.ComputeGlobalInterference()
• IpfcGlobalEvaluator.Assem
• IpfcGlobalEvaluator.Assem
• IpfcGlobalInterference.Volume
• IpfcGlobalInterference.SelParts
| Tocomputeall | theinterferenceswithin | an  | Assemblyonehasto | call |
| ------------ | ---------------------- | --- | ---------------- | ---- |
CMpfcInterference.CreateGlobalEvaluator()
witha
| IpfcAssemblyobjectasan |     | argument.This | callreturns |     |
| ---------------------- | --- | ------------- | ----------- | --- |
aIpfcGlobalEvaluatorobject.
ThepropertyIpfcGlobalEvaluator.Assem accessesthe assemblyto be
evaluated.
Themethod
| IpfcGlobalEvaluator.ComputeGlobalInterference() |                  |                    |     | determines |
| ----------------------------------------------- | ---------------- | ------------------ | --- | ---------- |
| thesetofall                                     | theinterferences | withintheassembly. |     |            |
Thismethodwillreturna sequenceof IpfcGlobalInterferenceobjectsor
nullif thereareno interferingparts. Eachobjectcontainsa pairof intersecting
partsandan objectrepresentingtheinterference volume,whichcanbe extracted
| byusing | IpfcGlobalInterference.SelPartsand |     |     |     |
| ------- | ---------------------------------- | --- | --- | --- |
IpfcGlobalInterference.Volumerespectively.
| Analyzing | Interference | Information |     |     |
| --------- | ------------ | ----------- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcSelectionPair.Create()
• CCpfcInterference.CreateSelectionEvaluator()
• IpfcSelectionEvaluator.Selections
• IpfcSelectionEvaluator.ComputeInterference()
• IpfcSelectionEvaluator.ComputeClearance()
• IpfcSelectionEvaluator.ComputeNearestCriticalDistance()
| ThemethodCCpfcSelectionPair.Create() |     |                  | createsa |           |
| ------------------------------------ | --- | ---------------- | -------- | --------- |
| IpfcSelectionPairobjectusing         |     | twoIpfcSelection |          | objectsas |
arguments.
| Areturnfromthis                               | methodwillserveas | an  | argumentto |           |
| --------------------------------------------- | ----------------- | --- | ---------- | --------- |
| CMpfcInterference.CreateSelectionEvaluator(), |                   |     |            | whichwill |
providea waytodeterminethe interferencedatabetweenthe twoselections.
| 230 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

IpfcSelectionEvaluator.Selectionswillextractand setthe objectto
beevaluated.
IpfcSelectionEvaluator.ComputeInterference()determinesthe
interferinginformationabout theprovidedselections.This methodwillreturnthe
IpfcInterferenceVolumeobjector nullif theselectionsdo no interfere.
IpfcSelectionEvaluator.ComputeClearance() computesthe
| clearancedatafor | thetwoselection.This | methodreturnsa |     |     |
| ---------------- | -------------------- | -------------- | --- | --- |
IpfcClearanceDataobject,whichcan
beused toobtainandsetclearance
distance,nearestpointsbetweenselections,anda booleanIsInterferening
variable.
IpfcSelectionEvaluator.ComputeNearestCriticalDistance()
| findsacriticalpoint | ofthedistance | functionbetween | twoselections. |     |
| ------------------- | ------------- | --------------- | -------------- | --- |
Thismethodreturns aIpfcCriticalDistanceData object,which isusedto
determineand setcritical points,surfaceparameters,and criticaldistancebetween
points.
| Analyzing | Interference | Volume |     |     |
| --------- | ------------ | ------ | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcInterferenceVolume.ComputeVolume()
• IpfcInterferenceVolume.Highlight()
• IpfcInterferenceVolume.Boundaries
| ThemethodIpfcInterferenceVolume.ComputeVolume() |                    |     |     | will |
| ----------------------------------------------- | ------------------ | --- | --- | ---- |
| calculatea valuefor                             | interferingvolume. |     |     |      |
ThemethodIpfcInterferenceVolume.Highlight() willhighlight the
interferingvolumewith thecolorprovided inthe argument tothe function.
ThepropertyIpfcInterferenceVolume.Boundaries willreturnasetof
| boundarysurfacedescriptorsfor |      | theinterferencevolume. |     |     |
| ----------------------------- | ---- | ---------------------- | --- | --- |
| Example                       | Code |                        |     |     |
Thesamplecodein thefilepfcGeometryExamples.vblocated at<creo_
| vbapi_loadpoint>/vbapi_examplesfinds |     |     | theinterferencein | an  |
| ------------------------------------ | --- | --- | ----------------- | --- |
assembly, highlightstheinterferingsurfaces,andhighlights calculatesthe
interferencevolume.
Thisapplicationfindsthe interferencein anassembly, highlightstheinterfering
| surfaces,andhighlightscalculatesthe |     | interferencevolume. |     |     |
| ----------------------------------- | --- | ------------------- | --- | --- |
GeometryEvaluation 231

18
Dimensions and Parameters
Overview................................................................................................................233
TheParamValueObject...........................................................................................233
ParameterObjects..................................................................................................234
DimensionObjects..................................................................................................240
Thischapterdescribes theVB APImethodsand classesthataffectdimensions and
parameters.
232 Creo®ParametricVBAPIUser’sGuide

Overview
DimensionsandparametersinCreo Parametrichavesimilarcharacteristicsbut
alsohavesignificantdifferences.In theVB API,thesimilaritiesbetween
dimensionsandparametersarecontainedin theIpfcBaseParameter class.
Thisclass allowsaccessto theparameter ordimension valueandto information
regardinga parameter'sdesignationandmodification.The differencesbetween
parametersanddimensions arerecognizablebecauseIpfcDimension inherits
fromtheclass IpfcModelItem, andcanbe assignedtolerances,whereas
parametersarenotIpfcModelItems andcannothavetolerances.
The ParamValue Object
Bothparametersanddimension objectscontainan objectoftype
IpfcParamValue. Thisobjectcontains theinteger,real, string,orBoolean
valueofthe parameteror dimension.Becauseofthe differentpossiblevaluetypes
thatcan be associatedwitha IpfcParamValueobjecttherearedifferent
methodsusedto accesseachvaluetypeandsomemethodswillnotbe applicable
forsomeIpfcParamValueobjects.Ifyoutryto usean incorrectmethodan
exceptionwillbethrown.
Accessing a ParamValue Object
MethodsandPropertyIntroduced:
• CMpfcModelItem.CreateIntParamValue()
• CMpfcModelItem.CreateDoubleParamValue()
• CMpfcModelItem.CreateStringParamValue()
• CMpfcModelItem.CreateBoolParamValue()
• CMpfcModelItem.CreateNoteParamValue()
• IpfcBaseParameter.Value
TheCMpfcModelItemutility class containsmethodsfor creatingeachtypeof
IpfcParamValueobject.Onceyouhaveestablishedthevaluetypein the
object,youcan changeit. The propertyIpfcBaseParameter.Value returns
theIpfcParamValueassociatedwith aparticular parameter ordimension.
ANoteParamValueisan integervalue thatrefers to theIDof aspecifiednote.
Tocreatea parameterofthis typetheidentifiednotemustalready existinthe
model.
Accessing the ParamValue Value
PropertiesIntroduced:
DimensionsandParameters 233

• IpfcParamValue.discr
• IpfcParamValue.IntValue
• IpfcParamValue.DoubleValue
• IpfcParamValue.StringValue
• IpfcParamValue.BoolValue
• IpfcParamValue.NoteId
ThepropertyIpfcParamValue.discr returnsaenumerationobjectthat
theIpfcParamValue
| identifiesthetypeofvalue |     | containedin |     | object.Use | this |
| ------------------------ | --- | ----------- | --- | ---------- | ---- |
informationwiththe specifiedpropertiestoaccess thevalue.Ifyouusean
incorrectpropertyanexception oftypeIpfcXBadGetParamValue willbe
thrown.
| Parameter |     | Objects |     |     |     |
| --------- | --- | ------- | --- | --- | --- |
Thefollowingsectionsdescribethe VBAPI methodsthataccessparameters. The
topicsareasfollows:
| • Creatingand             | AccessingParameterson |           | page234    |     |     |
| ------------------------- | --------------------- | --------- | ---------- | --- | --- |
| • ParameterSelection      |                       | Optionson | page235    |     |     |
| • ParameterInformationon  |                       | page237   |            |     |     |
| • ParameterRestrictionson |                       | page238   |            |     |     |
| Creating                  | and                   | Accessing | Parameters |     |     |
MethodsandPropertyIntroduced:
• IpfcParameterOwner.CreateParam()
• IpfcParameterOwner.CreateParamWithUnits()
• IpfcParameterOwner.GetParam()
• IpfcParameterOwner.ListParams()
• IpfcParameterOwner.SelectParam()
• IpfcParameterOwner.SelectParameters()
• IpfcFamColParam.RefParam
| IntheVB | API,models, | features,surfaces, | andedgesinheritfrom | the |     |
| ------- | ----------- | ------------------ | ------------------- | --- | --- |
IpfcParameterOwnerclass,becauseeachof theobjectscanbe assigned
| parametersin | Creo Parametric. |     |     |     |     |
| ------------ | ---------------- | --- | --- | --- | --- |
ThemethodIpfcParameterOwner.GetParam() getsa parametergivenits
name.
234 Creo®ParametricVBAPIUser’sGuide

| ThemethodIpfcParameterOwner.ListParams() |     |         |     | returnsa sequenceof |     |
| ---------------------------------------- | --- | ------- | --- | ------------------- | --- |
| allparametersassignedtothe               |     | object. |     |                     |     |
Tocreatea newparameterwith anameand aspecificvalue,call themethod
IpfcParameterOwner.CreateParam().
Tocreatea newparameterwith aname,aspecific value,andunits, callthe
methodIpfcParameterOwner.CreateParamWithUnits().
ThemethodIpfcParameterOwner.SelectParam() allows youto selecta
parameterfrom theCreo Parametricuserinterface.Thetopmodelfrom whichthe
| parametersareselectedmustbe                    |     | displayedinthe | currentwindow. |           |     |
| ---------------------------------------------- | --- | -------------- | -------------- | --------- | --- |
| ThemethodIpfcParameterOwner.SelectParameters() |     |                |                | allowsyou | to  |
interactivelyselectparametersfromtheCreo ParametricParameter dialogbox
| basedon theparameter                                | selectionoptions |     | specifiedbythe |          |     |
| --------------------------------------------------- | ---------------- | --- | -------------- | -------- | --- |
| IpfcParameterSelectionOptionsobject.Thetopmodelfrom |                  |     |                | whichthe |     |
parametersareselectedmustbe displayedinthe currentwindow.Referto the
sectionParameterSelectionOptions on page235for moreinformation.
ThepropertyIpfcFamColParam.RefParam returnsthe referenceparameter
| fromtheparameter          | columnin  | afamily  | table.            |       |     |
| ------------------------- | --------- | -------- | ----------------- | ----- | --- |
| Parameter                 | Selection | Options  |                   |       |     |
| Parameterselectionoptions |           | in theVB | APIarerepresented | bythe |     |
IpfcParameterSelectionOptionsclass.
MethodsandPropertiesIntroduced:
• CCpfcParameterSelectionOptions.Create
• IpfcParameterSelectionOptions.AllowContextSelection
• IpfcParameterSelectionOptions.Contexts
• .IpfcParameterSelectionOptions.AllowMultipleSelections
• IpfcParameterSelectionOptionsSelectButtonLabel
ThemethodCCpfcParameterSelectionOptions.Create createsanew
instanceof theIpfcParameterSelectionOptionsobjectthatis usedby
themethodIpfcParameterOwner.SelectParameters().
| Theparameterselectionoptions |     | areasfollows: |     |     |     |
| ---------------------------- | --- | ------------- | --- | --- | --- |
• AllowContextSelection—This booleanattributeindicates whetherto
allowparameter selectionfrommultiple contexts,or fromthe invoking
parameterowner. Bydefault,itis falseandallowsselectiononly fromthe
invokingparameterowner.Ifitis trueandif specificselection contextsarenot
| yetassigned,            | thenyoucanselecttheparametersfromanycontext. |     |     |     |     |
| ----------------------- | -------------------------------------------- | --- | --- | --- | --- |
| DimensionsandParameters |                                              |     |     |     | 235 |

| Usethe property |     |     |     |     |
| --------------- | --- | --- | --- | --- |
IpfcParameterSelectionOptions.AllowContextSelection
| tomodifythe | valueofthis | attribute. |     |     |
| ----------- | ----------- | ---------- | --- | --- |
• Contexts—The permittedparameter selectioncontexts inthe formof the
| IIpfcParameterSelectionContextsobject.Use |     |     |     | theproperty  |
| ----------------------------------------- | --- | --- | --- | ------------ |
| IpfcParameterSelectionOptions.Contexts    |     |     |     | to assignthe |
parameterselection context.Bydefault,youcan selectparametersfrom any
context.
| • Thetypesof                      | parameterselectioncontexts |             | areasfollows:               |            |
| --------------------------------- | -------------------------- | ----------- | --------------------------- | ---------- |
| ○ EpfcPARAMSELECT_MODEL—Specifies |                            |             | thatthetop                  | levelmodel |
| parameterscanbe                   |                            | selected.   |                             |            |
| ○ EpfcPARAMSELECT_PART—Specifies  |                            |             | thatany part’sparameters(at |            |
| anylevel                          | ofthetop                   | model)canbe | selected.                   |            |
○ EpfcPARAMSELECT_ASM—Specifies thatanyassembly’s parameters(at
| anylevel | ofthetop | model)canbe | selected. |     |
| -------- | -------- | ----------- | --------- | --- |
○ EpfcPARAMSELECT_FEATURE—Specifies thatanyfeature’s
| parameterscanbe |     | selected. |     |     |
| --------------- | --- | --------- | --- | --- |
○ EpfcPARAMSELECT_EDGE—Specifies thatany edge’s parameterscan
| be selected. |     |     |     |     |
| ------------ | --- | --- | --- | --- |
○ EpfcPARAMSELECT_SURFACE—Specifies thatanysurface’s
| parameterscanbe |     | selected. |     |     |
| --------------- | --- | --------- | --- | --- |
○ EpfcPARAMSELECT_QUILT—Specifies thatanyquilt’s parameterscan
| be selected. |     |     |     |     |
| ------------ | --- | --- | --- | --- |
○ EpfcPARAMSELECT_CURVE—Specifies thatanycurve’sparameters
| canbe | selected. |     |     |     |
| ----- | --------- | --- | --- | --- |
EpfcPARAMSELECT_COMPOSITE_CURVE—Specifies
| ○                                     |                 |               |                    | thatany |
| ------------------------------------- | --------------- | ------------- | ------------------ | ------- |
| compositecurve’s                      |                 | parameterscan | be selected.       |         |
| ○ EpfcPARAMSELECT_INHERITED—Specifies |                 |               | thatanyinheritance |         |
| feature’s                             | parameterscanbe | selected.     |                    |         |
| ○ EpfcPARAMSELECT_SKELETON—Specifies  |                 |               | thatanyskeleton’s  |         |
| parameterscanbe                       |                 | selected.     |                    |         |
| ○ EpfcPARAMSELECT_COMPONENT—Specifies |                 |               | thatanycomponent’s |         |
| parameterscanbe                       |                 | selected.     |                    |         |
• AllowMultipleSelections—This booleanattributeindicates whether
ornotto allowmultiple parametersto be selectedfromthe dialogbox,or only
asingleparameter. Bydefault,itis trueand allowsselectionofmultiple
parameters.
236 Creo®ParametricVBAPIUser’sGuide

Usethe property
IpfcParameterSelectionOptions.AllowMultipleSelec
tionsto modifythisattribute.
• SelectButtonLabel—The visiblelabelfor theselectbuttonin thedialog
box.
Usethe property
IpfcParameterSelectionOptionsSelectButtonLabel tosetthe
label.Ifnotset, thedefault labelin thelanguageofthe activeCreoParametric
sessionisdisplayed.
Parameter Information
MethodsandPropertiesIntroduced:
• IpfcBaseParameter.Value
• IpfcParameter.GetScaledValue()
• IpfcParameter.SetScaledValue()
• IpfcParameter.Units()
• IpfcBaseParameter.IsDesignated
• IpfcBaseParameter.IsModified
• IpfcBaseParameter.ResetFromBackup()
• IpfcParameter.Description
• IpfcParameter.GetRestriction()
• IpfcParameter.GetDriverType()
• IpfcParameter.Reorder()
• IpfcParameter.Delete()
• IpfcNamedModelItem.Name
Parametersinheritmethods fromtheIpfcBaseParameter,
IpfcParameterandIpfcNamedModelItem classes.
ThepropertyIpfcBaseParameter.Value returns thevalueof theparameter
ordimension.
ThemethodIpfcParameter.GetScaledValue() returns theparameter
valuein theunitsofthe parameter,insteadof theunitsofthe ownermodelas
returnedbyIpfcBaseParameter.Value.
ThemethodIpfcParameter.SetScaledValue() assignsthe parameter
valuein theunitsprovided,insteadofusing theunitsof theowner modelas
assumedby IpfcBaseParameter.Value.
DimensionsandParameters 237

ThemethodIpfcParameter.Units() returns theunitsassignedto the
parameter.
You canaccessthe designationstatus oftheparameter using theproperty
| IpfcBaseParameter.IsDesignated          |     |     | .                   |     |
| --------------------------------------- | --- | --- | ------------------- | --- |
| ThepropertyIpfcBaseParameter.IsModified |     |     | andthe method       |     |
| pfcBaseParameter.ResetFromBackup()      |     |     | enableyoutoidentify | a   |
modifiedparameteror dimension,andresetitto the laststoredvalue.Aparameter
issaidto be "modified"whenthe valuehasbeenchangedbutthe parameter's
| ownerhasnotyet                            | beenregenerated.            |     |                     |       |
| ----------------------------------------- | --------------------------- | --- | ------------------- | ----- |
| ThepropertyIpfcParameter.Description      |                             |     | returnstheparameter |       |
| description,ornull,                       | ifno descriptionisassigned. |     |                     |       |
| ThepropertyIpfcParameter.GetRestriction() |                             |     | identifies          | ifthe |
parameter’s valueis restrictedto acertainrangeor enumeration.It returnsthe
| IpfcParameterRestrictionobject.Referto |     |                  | thesectionParameter |     |
| -------------------------------------- | --- | ---------------- | ------------------- | --- |
| Restrictionsonpage238for               |     | moreinformation. |                     |     |
ThepropertyIpfcParameter.GetDriverType() returnsthedriver type
| foramaterial | parameter.The | drivertypesareasfollows: |     |     |
| ------------ | ------------- | ------------------------ | --- | --- |
• EpfcPARAMDRIVER_PARAM—Specifies thattheparametervalueisdriven
byanotherparameter.
• EpfcPARAMDRIVER_FUNCTION—Specifiesthatthe parametervalue is
| drivenby | afunction. |     |     |     |
| -------- | ---------- | --- | --- | --- |
• EpfcPARAMDRIVER_RELATION—Specifiesthatthe parametervalue is
| drivenby | arelation. Thisis | equivalenttothe | valueobtainedusing |     |
| -------- | ----------------- | --------------- | ------------------ | --- |
IpfcBaseParameter.IsRelationDriven for aparameterobject
type.
ThemethodIpfcParameter.Reorder() reordersthegiven parameterto
comeimmediatelyaftertheindicated parameterin theParameter dialogboxand
| informationfilesgeneratedbyCreo |     | Parametric. |                     |           |
| ------------------------------- | --- | ----------- | ------------------- | --------- |
| ThemethodIpfcParameter.Delete(  |     |             | permanentlyremovesa | specified |
parameter.
| ThepropertypfcModelItem.NamedModelItem.GetName |                     |     |     | accessesthe |
| ---------------------------------------------- | ------------------- | --- | --- | ----------- |
| nameofthe                                      | specifiedparameter. |     |     |             |
| Parameter                                      | Restrictions        |     |     |             |
CreoParametricallowsusers toassignspecifiedlimitationsto thevalueallowed
foragiven parameter(wherever theparameter appearsin themodel).Youcan
onlyread thedetailsof thepermittedrestrictionsfrom theVB API,butnot modify
thepermittedvaluesorrangeof values.Parameterrestrictionsin theVB APIare
| representedby | theclassIpfcParameterRestriction. |     |     |     |
| ------------- | --------------------------------- | --- | --- | --- |
238 Creo®ParametricVBAPIUser’sGuide

MethodIntroduced:
• IpfcParameterRestriction.Type
ThemethodIpfcParameterRestriction.Type returnsthe
IpfcRestrictionType objectcontainingthe typesofparameter restrictions.
Theparameterrestrictionsare ofthefollowingtypes:
• EpfcPARAMSELECT_ENUMERATION—Specifiesthattheparameteris
restrictedto alist ofpermittedvalues.
• EpfcPARAMSELECT_RANGE—Specifies thattheparameterislimitedto a
specifiedrangeofnumeric values.
Enumeration Restriction
TheEpfcPARAMSELECT_ENUMERATIONtypeof parameterrestrictionis
representedby theinterfaceIpfcParameterEnumeration. Itis achildof
theIpfcParameterRestriction interface.
PropertyIntroduced:
• ParameterEnumeration.PermittedValues
ThepropertyParameterEnumeration.PermittedValues returnsalist
ofpermittedparametervalues allowedby thisrestrictioninthe formof asequence
oftheIpfcParamValue objects.
Range Restriction
TheEpfcPARAMSELECT_RANGE typeofparameter restrictionisrepresented by
theinterfaceIpfcParameterRange.It isachild ofthe
IpfcParameterRestrictioninterface.
PropertiesIntroduced:
• IpfcParameterRange.Maximum
• IpfcParameterRange.Minimum
• IpfcParameterLimit.Type
• IpfcParameterLimit.Value
ThepropertyIpfcParameterRange.Maximum returnsthe maximumvalue
limitforthe parameterin theform ofthe IpfcParameterLimitobject.
ThepropertyIpfcParameterRange.Minimum returnsthe minimumvalue
limitforthe parameterin theform ofthe IpfcParameterLimitobject.
ThepropertyIpfcParameterLimit.Type returns the
IpfcParameterLimitTypecontainingthetypesof parameterlimits.The
parameterlimitsareof thefollowingtypes:
DimensionsandParameters 239

• EpfcPARAMLIMIT_LESS_THAN—Specifiesthatthe parametermustbe
lessthantheindicatedvalue.
• EpfcPARAMLIMIT_LESS_THAN_OR_EQUAL—Specifiesthatthe
| parametermustbe |     | lessthanorequalto |     | the indicatedvalue. |
| --------------- | --- | ----------------- | --- | ------------------- |
• EpfcPARAMLIMIT_GREATER_THAN—Specifiesthattheparametermust
| begreaterthanthe |     | indicatedvalue. |     |     |
| ---------------- | --- | --------------- | --- | --- |
• EpfcPARAMLIMIT_GREATER_THAN_OR_EQUAL—Specifies thatthe
| parametermustbe |     | greaterthanorequalto |     | the indicatedvalue. |
| --------------- | --- | -------------------- | --- | ------------------- |
ThepropertyIpfcParameterLimit.Value returnsthe boundaryvalueof
| theparameterlimit |       | inthe formof                               | theIpfcParamValueobject. |            |
| ----------------- | ----- | ------------------------------------------ | ------------------------ | ---------- |
| Example           | Code: | Updating                                   | Model                    | Parameters |
| Thesamplecodein   |       | thefilepfcDimensionAndParameterExamples.vb |                          |            |
locatedat<creo_vbapi_loadpoint>/vbapi_examples containsa
methodthatreadsa "properties"fileand createsorupdatesmodelparametersfor
eachproperty whichexistsinthe file.Sinceeachpropertyvalue isreturnedas a
String,autilitymethod parsesthe Stringintoint,double,or booleanvaluesif
possible.
| Dimension |     | Objects |     |     |
| --------- | --- | ------- | --- | --- |
DimensionobjectsincludestandardCreo Parametricdimensionsas wellas
referencedimensions.Dimension objectsenableyoutoaccess dimension
tolerancesandenableyouto setthevalue forthe dimension.Reference
| dimensionsallowneitherof |            | theseactions. |     |     |
| ------------------------ | ---------- | ------------- | --- | --- |
| Getting                  | Dimensions |               |     |     |
DimensionsandreferencedimensionsareCreo Parametricmodelitems.Seethe
sectionGettingModelItemObjectson page189formethods thatcanreturn
| IpfcDimensionandIpfcRefDimension |             |     |     | objects. |
| -------------------------------- | ----------- | --- | --- | -------- |
| Dimension                        | Information |     |     |          |
MethodsandPropertiesIntroduced:
• IpfcBaseParameter.Value
• IpfcBaseDimension.DimValue
• IpfcBaseParameter.IsDesignated
• IpfcBaseParameter.IsModified
• IpfcBaseParameter.ResetFromBackup()
240 Creo®ParametricVBAPIUser’sGuide

• IpfcBaseParameter.IsRelationDriven
• IpfcBaseDimension.DimType
• IpfcBaseDimension.Symbol
• IpfcBaseDimension.Texts
AlltheIpfcBaseParameter
|     |     |     | methodsareaccessible |     | to  | Dimensionsaswellas |
| --- | --- | --- | -------------------- | --- | --- | ------------------ |
Parameters.SeethesectionParameter Objectsonpage234for briefdescriptions.
Note
Youcannotsetthevalue ordesignationstatusofreferencedimension objects.
ThepropertyIpfcBaseDimension.DimValue accessesthedimension value
asadouble. Thispropertyprovidesa shortcutforaccessing thedimensions' values
| withoutusinga                                 | IpfcParamValueobject. |                                     |     |     |     |            |
| --------------------------------------------- | --------------------- | ----------------------------------- | --- | --- | --- | ---------- |
| ThepropertyIpfcBaseParameter.IsRelationDriven |                       |                                     |     |     |     | identifies |
| whetherthe                                    | partor                | assemblyrelationscontroladimension. |     |     |     |            |
ThepropertyIpfcBaseDimension.DimType returnsan enumerationobject
thatidentifies whetheradimensionis linear, radial,angular,or diametrical.
| ThepropertyIpfcBaseDimension.Symbol |     |             |           |         | returnsthe | dimensionor |
| ----------------------------------- | --- | ----------- | --------- | ------- | ---------- | ----------- |
| referencedimension                  |     | symbol(that | is,“d#”or | “rd#”). |            |             |
ThepropertyIpfcBaseDimension.Texts allows accessto thetextstrings
| thatprecede | orfollowthe | dimensionvalue. |     |     |     |     |
| ----------- | ----------- | --------------- | --- | --- | --- | --- |
| Dimension   |             | Tolerances      |     |     |     |     |
MethodsandPropertiesIntroduced:
• IpfcDimension.Tolerance
• CCpfcDimTolPlusMinus.Create()
• CCpfcDimTolSymmetric.Create()
• CCpfcDimTolLimits.Create()
• CCpfcDimTolSymSuperscript.Create()
• CCpfcDimTolISODIN.Create()
| Onlytrue                           | dimensionobjectscanhavegeometrictolerances. |                |     |              |               |           |
| ---------------------------------- | ------------------------------------------- | -------------- | --- | ------------ | ------------- | --------- |
| ThepropertyIpfcDimension.Tolerance |                                             |                |     |              | enablesyouto  | accessthe |
| dimensiontolerance.The             |                                             | objecttypesfor |     | thedimension | toleranceare: |           |
• IpfcDimTolLimits—Displaysdimension tolerancesasupperandlower
limits.
DimensionsandParameters 241

Note
This formatisnotavailablewhenonly thetolerance valuefor adimension
| is  | displayed. |     |     |     |     |     |     |     |
| --- | ---------- | --- | --- | --- | --- | --- | --- | --- |
• IpfcDimTolPlusMinus—Displays dimensionsas nominalwith plus-
| minustolerances.The |     |     | positiveandnegativevalues |     |     |     | areindependent. |     |
| ------------------- | --- | --- | ------------------------- | --- | --- | --- | --------------- | --- |
• IpfcDimTolSymmetric—Displays dimensionsas nominalwith asingle
| valuefor | boththe | positiveandthe |     |     | negativetolerance. |     |     |     |
| -------- | ------- | -------------- | --- | --- | ------------------ | --- | --- | --- |
• IpfcDimTolSymSuperscript—Displays dimensionsasnominalwith a
singlevalue forpositive andnegativetolerance.Thetextof thetoleranceis
| displayedina |     | superscriptformatwith |     |     | respectto |     | thedimensiontext. |     |
| ------------ | --- | --------------------- | --- | --- | --------- | --- | ----------------- | --- |
• IpfcDimTolISODIN—Displaysthe tolerancetabletype, tablecolumn,and
tablename,if thedimension toleranceissetto aholeorshaft table (DIN/ISO
standard).
| Anullvalue |     | issimilar | tothe | nominaloption |     | in CreoParametric. |     |     |
| ---------- | --- | --------- | ----- | ------------- | --- | ------------------ | --- | --- |
You candeterminewhetheragiventolerance isplus/minus,symmetric, limits,or
superscriptusing the followingexamplecode.If TypeOf(tolerance) Is
IpfcDimTolLimits.
| Example                                        | Code: | Setting                                    |     | Tolerances |                  | to  | a Specified | Range         |
| ---------------------------------------------- | ----- | ------------------------------------------ | --- | ---------- | ---------------- | --- | ----------- | ------------- |
| Thesamplecodein                                |       | thefilepfcDimensionAndParameterExamples.vb |     |            |                  |     |             |               |
| locatedat<creo_vbapi_loadpoint>/vbapi_examples |       |                                            |     |            |                  |     |             | showsautility |
| methodthatsetsangulartolerancesto              |       |                                            |     |            | aspecifiedrange. |     |             |               |
Theexamplecodeshowsautility methodthatsetsangulartolerancesto a
specifiedrange. First,the programdetermineswhetherthe dimensionpassedto it
isangular. Ifitis,themethodgets thedimensionvalue andaddsor subtractsthe
| rangetoitto | getthe | upperand |     | lowerlimits. |     |     |                                 |     |
| ----------- | ------ | -------- | --- | ------------ | --- | --- | ------------------------------- | --- |
| 242         |        |          |     |              |     |     | Creo®ParametricVBAPIUser’sGuide |     |

19
Relations
AccessingRelations................................................................................................244
AccessingPostRegenerationRelations....................................................................244
AddingaCustomizedFunctiontotheRelationsDialogBoxinCreo
Parametric ..........................................................................................................245
Thischapterdescribes howtoaccessrelationson allmodelsandmodelitems in
CreoParametricusing themethodsprovided in theVB API.
243

| Accessing | Relations |     |     |     |     |
| --------- | --------- | --- | --- | --- | --- |
IntheVB API,thesetof relationsonanymodelor modelitem isrepresented by
theIpfcRelationOwner interface.Models,features,surfaces,andedges
inheritfromthis interface,becauseeachobjectcanbe assignedrelationsin Creo
Parametric.
MethodsandPropertiesIntroduced:
• IpfcRelationOwner.RegenerateRelations()
• IpfcRelationOwner.DeleteRelations()
• IpfcRelationOwner.Relations
• IpfcRelationOwner.EvaluateExpression()
ThemethodIpfcRelationOwner.RegenerateRelations()
regeneratesthe relationsassignedto theowner item.It alsodetermineswhether
| thespecifiedrelationsetis                    |          | valid. |     |            |     |
| -------------------------------------------- | -------- | ------ | --- | ---------- | --- |
| ThemethodIpfcRelationOwner.DeleteRelations() |          |        |     | deletesall | the |
| relationsassignedto                          | theowner | item.  |     |            |     |
ThepropertyIpfcRelationOwner.Relations
|                                                 |          |                  | returnsthelist | ofinitial |     |
| ----------------------------------------------- | -------- | ---------------- | -------------- | --------- | --- |
| relationsassignedto                             | theowner | item asasequence | ofstrings.     |           |     |
| ThemethodIpfcRelationOwner.EvaluateExpression() |          |                  |                | evaluates |     |
thegivenrelations-basedexpression, andreturnstheresulting valuein the formof
theIpfcParamValueobject.Referto thesectionThe ParamValueObjecton
page233in thechapter DimensionsandParameters onpage232for more
| informationonthis | object.   |           |         |            |     |
| ----------------- | --------- | --------- | ------- | ---------- | --- |
| Example           | 1: Adding | Relations | between | Parameters | in  |
| a Solid           | Model     |           |         |            |     |
Thesamplecodein thefilepfcRelationsExamples.vblocatedat<creo_
vbapi_loadpoint>/vbapi_examplesdemonstrateshowto addrelations
| betweenparametersin | asolidmodel. |              |     |           |     |
| ------------------- | ------------ | ------------ | --- | --------- | --- |
| Accessing           | Post         | Regeneration |     | Relations |     |
MethodandPropertyIntroduced:
• IpfcModel.PostRegenerationRelations
• IpfcModel.RegeneratePostRegenerationRelations()
• IpfcModel.DeletePostRegenerationRelations()
ThepropertyIpfcModel.PostRegenerationRelations lists thepost-
regenerationrelationsassignedto themodel.It canbe NULL, ifnotset.
| 244 |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | ------------------------------- | --- | --- |

Note
Toworkwith post-regenerationrelations,usethepost-regeneration relations
| attributein                             | themethods |     |     |     |     |     |
| --------------------------------------- | ---------- | --- | --- | --- | --- | --- |
| IpfcRelationOwner.RegenerateRelations() |            |     |     |     | and |     |
IpfcRelationOwner.DeleteRelations().
You canregenerate therelationsetspost-regenerationin amodelusing themethod
IpfcModel.RegeneratePostRegenerationRelations().
Todeleteall thepost-regenerationrelationsin thespecifiedmodel, callthemethod
IpfcModel.DeletePostRegenerationRelations().
| Adding    | a Customized |     | Function |     | to the     |     |
| --------- | ------------ | --- | -------- | --- | ---------- | --- |
| Relations | Dialog       | Box | in Creo  |     | Parametric |     |
MethodsIntroduced:
• IpfcBaseSession.RegisterRelationFunction
| ThemethodIpfcBaseSession.RegisterRelationFunction |     |     |     |     |     | registers |
| ------------------------------------------------- | --- | --- | --- | --- | --- | --------- |
acustomfunctionthatisincludedin thefunction listof theRelations dialogbox
inCreo Parametric.You canaddthe customfunction to relationsthatare addedto
models,features,orotherrelationowners. Theregistrationmethodtakesthe
followinginputarguments:
| • Name—Thename                     | ofthecustomfunction. |     |     |                   |     |         |
| ---------------------------------- | -------------------- | --- | --- | ----------------- | --- | ------- |
| • IpfcRelationFunctionOptions—This |                      |     |     | objectcontainsthe |     | options |
thatdeterminethe behaviorof thecustomrelationfunction.Referto the
| sectionRelationFunctionOptions      |     | on  | page246for |                         | moreinformation. |     |
| ----------------------------------- | --- | --- | ---------- | ----------------------- | ---------------- | --- |
| • IpfcRelationFunctionListener—This |     |     |            | objectcontainstheaction |                  |     |
listenermethodsfor theimplementation ofthecustomfunction.Referto the
sectionRelationFunctionListeners on page247formoreinformation.
Relations 245

Note
The VBAPI relationfunctions arevalid onlywhen thecustomfunctionhas
beenregisteredbythe application.Ifthe applicationis notrunning ornot
present,modelsthatcontain user-definedrelationscannotevaluatethese
relations. Inthissituation,therelationsaremarkedas errors.However, these
errorscan be commenteduntilneededatalater timewhen therelations
|          | functions | arereactivatedin | aCreo   | Parametricsession. |     |     |
| -------- | --------- | ---------------- | ------- | ------------------ | --- | --- |
| Relation |           | Function         | Options |                    |     |     |
MethodsandPropertiesIntroduced:
• CCpfcRelationFunctionOptions.Create()
• IpfcRelationFunctionOptions.ArgumentTypes
• CCpfcRelationFunctionArgument.Create()
• IpfcRelationFunctionArgument.Type
• .IpfcRelationFunctionArgument.IsOptional
• IpfcRelationFunctionOptions.EnableTypeChecking
• IpfcRelationFunctionOptions.EnableArgumentCheckMethod
• IpfcRelationFunctionOptions.EnableExpressionEvaluationMethod
• IpfcRelationFunctionOptions.EnableValueAssignmentMethod
Usethe methodCCpfcRelationFunctionOptions.Create() tocreate
theIpfcRelationFunctionOptionsobjectcontainingtheoptions to
enableor disablevariousrelationfunction relatedfeatures.Usethe methodslisted
| abovetoaccess |                                     | andmodifytheoptions. |     |     | Theseoptionsare | asfollows: |
| ------------- | ----------------------------------- | -------------------- | --- | --- | --------------- | ---------- |
| •             | ArgumentTypes—Thetypesofargumentsin |                      |     |     | theform         | ofthe      |
IpfcRelationFunctionArgument object.Bydefault,this parameteris
|     | null,indicatingthatno                               |     | arguments | arepermitted. |     |     |
| --- | --------------------------------------------------- | --- | --------- | ------------- | --- | --- |
|     | Usethe methodCCpfcRelationFunctionArgument.Create() |     |           |               |     | to  |
createtheIpfcRelationFunctionArgument
objectcontainingthe
|     | attributesof       | theargumentspassedtothe |            |     | customrelationfunction. |     |
| --- | ------------------ | ----------------------- | ---------- | --- | ----------------------- | --- |
|     | Theseattributesare |                         | asfollows: |     |                         |     |
○ Type—Thetypeofargumentvaluesuch asdouble,integer, andsoon in
|     | theform | ofthe | IpfcParamValueTypeobject. |     |     |     |
| --- | ------- | ----- | ------------------------- | --- | --- | --- |
○ IsOptional—This booleanattributespecifies whethertheargumentis
|     | optional, | indicatingthatitcan |     | beskipped | when                            | acallto thecustom |
| --- | --------- | ------------------- | --- | --------- | ------------------------------- | ----------------- |
| 246 |           |                     |     |           | Creo®ParametricVBAPIUser’sGuide |                   |

relationfunctionis made.Theoptionalargumentsmustfallatthe endof
| theargumentlist.Bydefault,this |     |     | attributeisfalse. |
| ------------------------------ | --- | --- | ----------------- |
• EnableTypeChecking—This booleanattributedetermineswhetherornotto
checktheargumenttypesinternally. Bydefault,itisfalse. Ifthis attributeis
settofalse, Creo Parametricdoesnotneedto knowthecontents ofthe
arguments array.The customfunction musthandleallusererrors insuch a
situation.
• EnableArgumentCheckMethod—This booleanattributedetermineswhetheror
notto enabletheargumentschecklistener function.Bydefault,itis false.
• EnableExpressionEvaluationMethod—This booleanattributedetermines
whetherornotto enabletheevaluatelistener function.Bydefault,itistrue.
• EnableValueAssignmentMethod—Thisbooleanattribute determineswhether
ornotto enablethevalue assignmentlistener function.Bydefault,itis false.
| Relation | Function | Listeners |     |
| -------- | -------- | --------- | --- |
TheinterfaceIpfcRelationFunctionListener providesthemethod
| signaturesto | implementacustomrelationfunction. |     |     |
| ------------ | --------------------------------- | --- | --- |
MethodsIntroduced:
• IpfcRelationFunctionListener.CheckArguments
• IpfcRelationFunctionListener.AssignValue
• IpfcRelationFunctionListener.EvaluateFunction
ThemethodIpfcRelationFunctionListener.CheckArguments
checksthevalidityof theargumentspassedto thecustomfunction.This listener
methodtakesthefollowinginputarguments:
| • Theowner          | ofthe relationbeingevaluated |             |                |
| ------------------- | ---------------------------- | ----------- | -------------- |
| • Thecustomfunction |                              | name        |                |
| • Asequenceof       | arguments                    | passedtothe | customfunction |
Iftheimplementationof thismethoddeterminesthattheargumentsare notvalid
forthecustomfunction, thenthe listenermethodreturns false.Otherwise,it
returnstrue.
ThemethodIpfcRelationFunctionListener.EvaluateFunction
evaluatesacustomrelationfunctioninvokedon therighthand sideof arelation.
| Thislistener        | methodtakesthefollowinginputarguments: |             |                |
| ------------------- | -------------------------------------- | ----------- | -------------- |
| • Theowner          | ofthe relationbeingevaluated           |             |                |
| • Thecustomfunction |                                        | name        |                |
| • Asequenceof       | arguments                              | passedtothe | customfunction |
Relations 247

| You mustreturn | thecomputedresult |     | ofthecustomrelationfunction. |     |     |     |
| -------------- | ----------------- | --- | ---------------------------- | --- | --- | --- |
ThemethodIpfcRelationFunctionListener.AssignValue
evaluates
acustomrelationfunction invokedon thelefthandside ofarelation. Itallows you
toinitialize propertiestobe storedandused byyour application.Thislistener
methodtakesthefollowinginputarguments:
| • Theowner          | ofthe | relationbeingevaluated |             |                |     |     |
| ------------------- | ----- | ---------------------- | ----------- | -------------- | --- | --- |
| • Thecustomfunction |       | name                   |             |                |     |     |
| • Asequenceof       |       | arguments              | passedtothe | customfunction |     |     |
• Thevalue obtainedbyCreo Parametricfromevaluating therighthandside of
therelation
| Example  | 2:       | Adding | and Implementing |     | a New | Custom |
| -------- | -------- | ------ | ---------------- | --- | ----- | ------ |
| Relation | Function |        |                  |     |       |        |
Thesamplecodein thefilepfcRelationsExamples.vblocatedat<creo_
| vbapi_loadpoint>/vbapi_exampleshasthefunction |     |     |     |     | addRelation |     |
| --------------------------------------------- | --- | --- | --- | --- | ----------- | --- |
function,whichdefinesthe optionsfor anewcustomrelationfunction and
registersitinthe currentsession. The RelationListenerclass containsthe
| CheckArguments,             |     | AssignValueandEvaluateFunction |                        |                                 |         | listener |
| --------------------------- | --- | ------------------------------ | ---------------------- | ------------------------------- | ------- | -------- |
| methodsthatarecalledwhenthe |     |                                | customrelationfunction |                                 | isused. |          |
| 248                         |     |                                |                        | Creo®ParametricVBAPIUser’sGuide |         |          |

20
Assemblies and Components
StructureofAssembliesandAssemblyObjects.........................................................250
AssemblingComponents.........................................................................................255
RedefiningandReroutingAssemblyComponents.....................................................259
ExplodedAssemblies..............................................................................................259
SkeletonModels......................................................................................................260
Thischapterdescribes theVB APIfunctions thataccessthefunctions ofaCreo
Parametricassembly. You mustbefamiliarwith thefollowingbeforeyouread this
section:
• TheSelection Object
• CoordinateSystems
• TheGeometrysection
249

Structure of Assemblies and Assembly
Objects
TheobjectIpfcAssemblyis aninstance ofIpfcSolid. The
IpfcAssemblyobjectcantherefore beused asinputto anyof theIpfcSolid
andIpfcModelmethods applicableto assemblies.Howeverassemblies donot
containsolidgeometryitems.The onlygeometryin theassembly isdatums
(points,planes,axes,coordinatesystems,curves,andsurfaces).Thereforesolid
assemblyfeaturessuchas holesandslotswillnotcontain activesurfacesor edges
inthe assemblymodel.
Thesolidgeometryofan assembly iscontainedin itscomponents. Acomponent
isafeatureof typeIpfcComponentFeat, whichisa referencetoa partor
anotherassembly, andasetofparametric constraintsfor determiningits
geometricallocationwithin theparent assembly.
Assemblyfeaturesthataresolid,such asholesandslots,andtherefore affectthe
solidgeometryof partsin the assemblyhierarchy,do notthemselvescontainthe
geometryitemsthatdescribethosemodifications.Theseitems arealways
containedin theparts whosegeometryis modified,within local featurescreated
forthatpurpose.
Theimportantfunctionsfor assembliesarethosethatoperateon thecomponents
ofan assembly. TheobjectIpfcComponentFeat, whichisan instance of
IpfcFeatureis definedforthatpurpose.Eachassembly componentistreated
asavarietyoffeature,andthe integeridentifier ofthecomponentisalso the
featureidentifier.
Anassemblycan containa hierarchyof assembliesandparts atmanylevels,in
whichsomeassemblies andparts mayappear morethanonce.Toidentify therole
ofanydatabase itemin thecontext oftheroot assembly, itisnotsufficientto have
theintegeridentifier oftheitem andthehandleto itsowningpart orassembly,as
wouldbe providedby itsIpfcFeature description.
Itisalso necessaryto givethefull pathof theassembly-componentreferences
downfromthe rootassembly tothe partor assemblythatownsthe databaseitem.
Thisisthe purposeof theobjectIpfcComponentPath, whichisusedas the
inputtothe VBAPI assemblyfunctions.
Thefollowingfigureshowsan assemblyhierarchywith twoexamplesofthe
contentsofa IpfcComponentPath object.
250 Creo®ParametricVBAPIUser’sGuide

Intheassembly shownin thefigure,subassembly Ciscomponentidentifier 11
withinassembly A,Part Bis componentidentifier 3within assemblyAB, andso
on.Thesubassembly AB occurstwice.Torefer to thetwooccurrences ofpart B,
usethe following:
| (?)Component               | B’  | Component            | B"  |               |
| -------------------------- | --- | -------------------- | --- | ------------- |
| ComponentIds.Item(0)       | = 2 | ComponentIds.Item(1) |     | = 11          |
| ComponentIds.Item(1)       | = 2 | ComponentIds.Item(2) |     | = 6           |
| ComponentIds.Item(2)       | = 5 | ComponentIds.Item(3) |     | = 12          |
| ComponentIds.Item(3)       | = 2 | ComponentIds.Item(4) |     | = 3           |
| ComponentIds.Item(4)       | = 3 |                      |     |               |
| TheobjectIpfcComponentPath |     | isone ofthemain      |     | portionsofthe |
IpfcSelectionobject.
| Assembly | Components |     |     |     |
| -------- | ---------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcComponentFeat.IsBulkitem
• IpfcComponentFeat.GetIsSubstitute
• IpfcComponentFeat.CompType
• IpfcComponentFeat.ModelDescr
AssembliesandComponents 251

• IpfcComponentFeat.IsPlaced
• IpfcComponentFeat.IsPackaged
• IpfcComponentFeat.IsUnderconstrained
• IpfcComponentFeat.IsFrozen
• IpfcComponentFeat.Position
• IpfcComponentFeat.CopyTemplateContents()
• IpfcComponentFeat.CreateReplaceOp()
ThepropertyIpfcComponentFeat.IsBulkitem
|     |     |     |     |     |     |     | identifies | whetheran |
| --- | --- | --- | --- | --- | --- | --- | ---------- | --------- |
assemblycomponentisabulk item.Abulk itemis anon-geometricassembly
| featurethatshould                            |     | appearin | an  | assemblybillofmaterials. |     |     |     |              |
| -------------------------------------------- | --- | -------- | --- | ------------------------ | --- | --- | --- | ------------ |
| ThepropertyIpfcComponentFeat.GetIsSubstitute |     |          |     |                          |     |     |     | returnsatrue |
valueifthe componentis substituted,else itreturnsafalse. Whenyousubstitutea
componentin asimplifiedrepresentation,youtemporarily excludethesubstituted
| componentandsuperimpose               |     |     | thesubstitutingcomponentin |     |     |            | itsplace. |            |
| ------------------------------------- | --- | --- | -------------------------- | --- | --- | ---------- | --------- | ---------- |
| ThepropertyIpfcComponentFeat.CompType |     |     |                            |     |     | returnsthe |           | typeof the |
assemblycomponent.The componenttypeidentifiesthepurpose ofthe
| componentin                             | amanufacturing     |     |     | assembly.    |     |     |            |       |
| --------------------------------------- | ------------------ | --- | --- | ------------ | --- | --- | ---------- | ----- |
| ThepropertyIpfcComponentFeat.ModelDescr |                    |     |     |              |     |     | returnsthe | model |
| descriptorof                            | thecomponentpartor |     |     | subassembly. |     |     |            |       |
Note
| From                         | Pro/ENGINEERWildfire4.0onwards,theproperty |     |     |            |        |                        |     |     |
| ---------------------------- | ------------------------------------------ | --- | --- | ---------- | ------ | ---------------------- | --- | --- |
| IpfcComponentFeat.ModelDescr |                                            |     |     |            | throws | anexception            |     |     |
| IpfcXtoolkitCantOpen         |                                            |     |     | ifcalledon | an     | assemblycomponentwhose |     |     |
immediategenericisnot insession. Handle thisexceptionand typecastthe
| assembly                                   | componentas |     | IpfcSolid,whichin |        |     | turncan | betypecast | as        |
| ------------------------------------------ | ----------- | --- | ----------------- | ------ | --- | ------- | ---------- | --------- |
| IpfcFamilyMember,andusethe                 |             |     |                   | method |     |         |            |           |
| IpfcFamilyMember.GetImmediateGenericInfo() |             |     |                   |        |     |         |            | toget the |
modeldescriptoroftheimmediategenericmodel. Ifyouwishto switchoff
thisbehavior andcontinueto runlegacyapplicationsin thepre-Wildfire4.0
| mode,settheconfigurationoption        |         |     |                            | retrieve_instance_ |     |                                 |     |     |
| ------------------------------------- | ------- | --- | -------------------------- | ------------------ | --- | ------------------------------- | --- | --- |
| dependenciesto                        |         |     | instance_and_generic_deps. |                    |     |                                 |     |     |
| ThepropertyIpfcComponentFeat.IsPlaced |         |     |                            |                    |     | determineswhetherthe            |     |     |
| componentis                           | placed. |     |                            |                    |     |                                 |     |     |
| 252                                   |         |     |                            |                    |     | Creo®ParametricVBAPIUser’sGuide |     |     |

Note
Once acomponentisconstrainedorpackaged, itcannotbe made unplaced
again.
Acomponentof anassembly thatiseitherpartiallyconstrainedorunconstrained
| isknown | asapackagedcomponent. |     | Usethe | property |     |
| ------- | --------------------- | --- | ------ | -------- | --- |
IpfcComponentFeat.IsPackagedto determineif thespecifiedcomponent
ispackaged.
ThepropertyIpfcComponentFeat.IsUnderconstrained determinesif
thespecifiedcomponentis underconstrained,thatis,itpossessessomeconstraints
butisnotfullyconstrained.
ThepropertyIpfcComponentFeat.IsFrozen determinesifthe specified
componentis frozen.Thefrozencomponentbehavessimilarto thepackaged
| componentanddoes |     | notfollowtheconstraints |     | thatyouspecify. |     |
| ---------------- | --- | ----------------------- | --- | --------------- | --- |
ThepropertyIpfcComponentFeat.Position retrieves thecomponent’s
initialposition beforeconstraintsandmovementshavebeenapplied.Ifthe
componentis packagedthis positionis thesameasthe constraint’s actualposition.
Thispropertymodifies theassembly componentdatabut doesnotregenerate the
| assemblycomponent.Toregenerate |     |     | thecomponent, |     | usethe method |
| ------------------------------ | --- | --- | ------------- | --- | ------------- |
IpfcComponentFeat.Regenerate.
ThemethodIpfcComponentFeat.CopyTemplateContents() copies
| thetemplatemodelintothemodelof |     |     | thespecifiedcomponent. |     |     |
| ------------------------------ | --- | --- | ---------------------- | --- | --- |
ThemethodIpfcCompontentFeat.CreateReplaceOp() createsa
replacementoperationusedto swapacomponentautomaticallywith arelated
| component.The | replacementoperationcanbe |     |     | usedasan | argumentto |
| ------------- | ------------------------- | --- | --- | -------- | ---------- |
IpfcSolid.ExecuteFeatureOps.
| Example | Code: | Replacing | Instances |     |     |
| ------- | ----- | --------- | --------- | --- | --- |
thefilepfcAssembliesExamples.vblocated
Thesamplecodein at
<creo_vbapi_loadpoint>/vbapi_examplescontains asingle static
utilitymethod.This methodtakesan assemblyfor an argument.It searches
throughthe assemblyfor allcomponentsthatareinstances ofthemodel"bolt". It
| thenreplaces | allsuchoccurrences |             | with | adifferentinstance | ofbolt. |
| ------------ | ------------------ | ----------- | ---- | ------------------ | ------- |
| Regenerating |                    | an Assembly |      | Component          |         |
MethodIntroduced:
• IpfcComponentFeat.Regenerate()
AssembliesandComponents 253

ThemethodIpfcComponentFeat.Regenerate() regeneratesan assembly
component.The methodregeneratesthe assemblycomponentjustas in an
interactiveCreo Parametricsession.
Creating a Component Path
MethodsIntroduced
• CMpfcAssembly.CreateComponentPath()
ThemethodCMpfcAssembly.CreateComponentPath() returnsa
componentpathobject,giventheAssemblymodelandthe integerid pathtothe
desiredcomponent.
Component Path Information
MethodsandPropertiesIntroduced:
• IpfcComponentPath.Root
• IpfcComponentPath.ComponentIds
• IpfcComponentPath.Leaf
• IpfcComponentPath.Transform()
• IpfcComponentPath.SetTransform()
• IpfcComponentPath.IsVisible()
ThepropertyIpfcComponentPath.Root returnsthe assemblyattheheadof
thecomponentpathobject.
ThepropertyIpfcComponentPath.ComponentIds returns thesequenceof
idswhichis thepathto theparticular component.
ThepropertyIpfcComponentPath.Leaf returnsthe solidmodelattheend
ofthecomponentpath.
ThemethodIpfcComponentPath.Transform() returnsthecoordinate
systemtransformationbetweenthe assemblyand theparticularcomponent. Ithas
anoption to providethetransformationfrom bottomto top,orfrom topto bottom.
Thismethoddescribes thecurrentpositionandthe orientationofthe assembly
componentin theroot assembly.
ThemethodIpfcComponentPath.SetTransform() appliesa temporary
transformationtothe assemblycomponent, similarto thetransformationthattakes
placein an explodedstate.The transformationwillonlybe appliedif theassembly
isusing DynamicPositioning.
ThemethodIpfcComponentPath.IsVisible() identifiesif aparticular
componentis visibleinany simplifiedrepresentation.
254 Creo®ParametricVBAPIUser’sGuide

| Assembling |     | Components |     |     |     |
| ---------- | --- | ---------- | --- | --- | --- |
MethodsIntroduced:
• IpfcAssembly.AssembleComponent()
• IpfcAssembly.AssembleByCopy()
• IpfcComponentFeat.GetConstraints()
• IpfcComponentFeat.SetConstraints()
• IpfcComponentFeat.GetConstraintsWithCompPath()
| ThemethodIpfcAssembly.AssembleComponent() |     |     |     |     | adds aspecified |
| ----------------------------------------- | --- | --- | --- | --- | --------------- |
componentmodelto theassembly atthespecifiedinitialposition.The positionis
specifiedin theformatdefinedby theIpfcTransform3D. Specifythe
orientationofthethree axesandthe positionofthe originofthe component
coordinatesystem,with respectto thetargetassembly coordinatesystem.
Note
Ifthe transformmatrixpassedas theinitialpositionofthe componentis
| incorrectandnon-orthonormal,     |     | themethod |     |            |       |
| -------------------------------- | --- | --------- | --- | ---------- | ----- |
| IpfcAssembly.AssembleComponent() |     |           |     | returnsthe | error |
pfcXToolkitBadInputs.In
suchscenario,youcanusethemethod
| CMpfcBase.MakeMatrixOrthonormal()      |     |                       |     | toconvertthis        | non- |
| -------------------------------------- | --- | --------------------- | --- | -------------------- | ---- |
| orthonormalmatrixto                    |     | an orthonormalmatrix. |     |                      |      |
| ThemethodIpfcAssembly.AssembleByCopy() |     |                       |     | createsanewcomponent |      |
inthe specifiedassembly by copyingfromthe specifiedcomponent. Ifno modelis
specified,thenthenewcomponentiscreated empty.The inputparametersfor this
methodare:
• LeaveUnplaced—Iftrue thecomponentisunplaced. Iffalsethecomponentis
placedatadefaultlocationin theassembly.Unplaced componentsbelongto
anassembly withoutbeingassembledor packaged.These componentsappear
inthe modeltree,butnotin the graphicwindow.Unplaced componentscanbe
| constrainedor | packagedbyselecting |     | themfrom | themodeltreefor |     |
| ------------- | ------------------- | --- | -------- | --------------- | --- |
redefinition.Whenitsparent assemblyis retrievedintomemory,an unplaced
| componentis                                 | alsoretrieved. |              |                       |     |              |
| ------------------------------------------- | -------------- | ------------ | --------------------- | --- | ------------ |
| • ModelToCopy—Specify                       |                | themodeltobe | copiedintotheassembly |     |              |
| • NewModelName—Specifyaname                 |                |              | forthe copiedmodel    |     |              |
| ThemethodIpfcComponentFeat.GetConstraints() |                |              |                       |     | retrievesthe |
| constraintsfora                             | givenassembly  | component.   |                       |     |              |
AssembliesandComponents 255

| ThemethodIpfcComponentFeat.SetConstraints() |     |     | allowsyouto | set |
| ------------------------------------------- | --- | --- | ----------- | --- |
theconstraintsfor aspecifiedassemblycomponent. Theinput parametersfor this
methodare:
• Constraints—Constraintsfor theassembly component.These constraintsare
| explainedin | detailin thelatersections. |     |     |     |
| ----------- | -------------------------- | --- | --- | --- |
• ReferenceAssembly—Thepathto theowner assembly, iftheconstraints have
| externalreferencesto | othermembers | ofthe toplevelassembly.Ifthe |     |     |
| -------------------- | ------------ | ---------------------------- | --- | --- |
constraintsareapplied onlyto theassembly componentthenthe valueofthis
| parametershouldbe | null. |     |     |     |
| ----------------- | ----- | --- | --- | --- |
Thismethodmodifies thecomponentfeaturedataandregeneratestheassembly
component.
ThemethodIpfcComponentFeat.GetConstraintsWithCompPath()
retrievestheconstraints fora givenassemblycomponentusing theinputargument
CompPathwhichis thepathto theowner assembly. Passthis inputargument
CompPath,if theconstraintshavereferencesto othermembersof thetoplevel
assembly. Passitas Null,if theconstraintshavereferencesonly tothe owner
assembly.
| Constraint | Attributes |     |     |     |
| ---------- | ---------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcConstraintAttributes.Create()
• IpfcConstraintAttributes.Force
• IpfcConstraintAttributes.Ignore
| ThemethodCCpfcConstraintAttributes.Create() |     |     | returnsthe |     |
| ------------------------------------------- | --- | --- | ---------- | --- |
constraintattributesobjectbasedon thevaluesof thefollowinginputparameters:
• Ignore—Constraintisignoredduringregeneration.Use thiscapability tostore
extraconstraintson thecomponent, whichallowsyouto quicklytoggle
betweendifferentconstraints.
| • Force—Constrainthasto | beforcedfor       | lineandpointalignment. |     |     |
| ----------------------- | ----------------- | ---------------------- | --- | --- |
| • None—Noconstraint     | attributes.Thisis | thedefaultvalue.       |     |     |
| Assembling              | a Component       | Parametrically         |     |     |
You canpositiona componentrelativeto itsneighbors(componentsor assembly
features)sothatitspositionis updatedasitsneighbors moveorchange. Thisis
calledparametricassembly. Creo Parametricallowsyouto specifyconstraints to
determinehowandwherethe componentrelatesto theassembly.You canaddas
manyconstraintsasyouneedto makesurethatthe assemblymeets thedesign
intent.
256 Creo®ParametricVBAPIUser’sGuide

MethodsandPropertiesIntroduced:
• CCpfcComponentConstraint.Create()
• IpfcComponentConstraint.Type
• IpfcComponentConstraint.AssemblyReference
• IpfcComponentConstraint.AssemblyDatumSide
• IpfcComponentConstraint.ComponentReference
• IpfcComponentConstraint.ComponentDatumSide
• IpfcComponentConstraint.Offset
• IpfcComponentConstraint.Attributes
• IpfcComponentConstraint.UserDefinedData
| ThemethodCCpfcComponentConstraint.Create() |     |              |     |                         |     |     | returnsthe |     |
| ------------------------------------------ | --- | ------------ | --- | ----------------------- | --- | --- | ---------- | --- |
| componentconstraint                        |     | objecthaving |     | thefollowingparameters: |     |     |            |     |
• EpfcComponentConstraintType—UsingtheTYPEoptions,you canspecify the
|     | placementconstrainttypes. |     |     | Theyareasfollows: |     |     |     |     |
| --- | ------------------------- | --- | --- | ----------------- | --- | --- | --- | --- |
EpfcASM_CONSTRAINT_MATE—Use
|     | ○                                 |     |        |                     | this       | optionto   | maketwosurfaces |           |
| --- | --------------------------------- | --- | ------ | ------------------- | ---------- | ---------- | --------------- | --------- |
|     | touchoneanother,                  |     | thatis | coincidentandfacing |            | eachother. |                 |           |
|     | ○ EpfcASM_CONSTRAINT_MATE_OFF—Use |     |        |                     |            | this       | optionto        | maketwo   |
|     | planar surfacesparallelandfacing  |     |        | eachother.          |            |            |                 |           |
|     | ○ EpfcASM_CONSTRAINT_ALIGN—Use    |     |        |                     | thisoption |            | tomake          | twoplanes |
coplanar, twoaxescoaxial andtwopointscoincident.You canalso align
|     | revolvedsurfaces                   |     | oredges. |         |     |               |     |                |
| --- | ---------------------------------- | --- | -------- | ------- | --- | ------------- | --- | -------------- |
|     | ○ EpfcASM_CONSTRAINT_ALIGN_OFF—Use |     |          |         |     | thisoption    |     | to aligntwo    |
|     | planar surfacesatan                |     |          | offset. |     |               |     |                |
|     | ○ EpfcASM_CONSTRAINT_INSERT—Use    |     |          |         |     | this optionto |     | inserta "male" |
revolvedsurface intoa``female''revolvedsurface,making theirrespective
axescoaxial.
|     | ○ EpfcASM_CONSTRAINT_ORIENT—Use |               |     |                   |     | this optionto |     | maketwoplanar |
| --- | ------------------------------- | ------------- | --- | ----------------- | --- | ------------- | --- | ------------- |
|     | surfacesto                      | be parallelin |     | thesamedirection. |     |               |     |               |
EpfcASM_CONSTRAINT_CSYS—Use
|     | ○   |     |     |     | this | optionto | placea | component |
| --- | --- | --- | --- | --- | ---- | -------- | ------ | --------- |
in anassembly by aligningthe coordinatesystem ofthe componentwith
|     | thecoordinatesystemof               |                      |       | theassembly. |     |            |               |               |
| --- | ----------------------------------- | -------------------- | ----- | ------------ | --- | ---------- | ------------- | ------------- |
|     | ○ EpfcASM_CONSTRAINT_TANGENT—Use    |                      |       |              |     | thisoption |               | to controlthe |
|     | contact                             | oftwosurfacesattheir |       | tangents.    |     |            |               |               |
|     | ○ EpfcASM_CONSTRAINT_PNT_ON_SRF—Use |                      |       |              |     |            | this optionto | controlthe    |
|     | contact                             | ofasurface           | witha | point.       |     |            |               |               |
AssembliesandComponents 257

|     | ○ EpfcASM_CONSTRAINT_EDGE_ON_SRF—Use   |     |             |     |                |     | thisoption |            | to control |
| --- | -------------------------------------- | --- | ----------- | --- | -------------- | --- | ---------- | ---------- | ---------- |
|     | thecontact                             | ofa | surfacewith |     | astraightedge. |     |            |            |            |
|     | ○ EpfcASM_CONSTRAINT_DEF_PLACEMENT—Use |     |             |     |                |     |            | thisoption | to align   |
thedefaultcoordinatesystem ofthecomponentto thedefaultcoordinate
|     | systemof                            | theassembly. |     |     |     |     |      |          |            |
| --- | ----------------------------------- | ------------ | --- | --- | --- | --- | ---- | -------- | ---------- |
|     | ○ EpfcASM_CONSTRAINT_SUBSTITUTE—Use |              |     |     |     |     | this | optionin | simplified |
representationswhen acomponenthasbeensubstituted withsomeother
model
|     | ○ EpfcASM_CONSTRAINT_PNT_ON_LINE—Use |                             |          |         |     |            | thisoption |     | to control |
| --- | ------------------------------------ | --------------------------- | -------- | ------- | --- | ---------- | ---------- | --- | ---------- |
|     | thecontact                           | ofa                         | linewith | apoint. |     |            |            |     |            |
|     | ○ EpfcASM_CONSTRAINT_FIX—Use         |                             |          |         |     | thisoption | toforcethe |     | component  |
|     | to remainin                          | itscurrentpackagedposition. |          |         |     |            |            |     |            |
○ EpfcASM_CONSTRAINT_AUTO—Use this optionin theuserinterfaceto
allowan automaticchoiceofconstraint typebasedupon thereferences.
○ EpfcASM_CONSTRAINT_EXPLICIT
| •   | AssemblyReference—Areferencein |     |     |     | theassembly. |     |     |     |     |
| --- | ------------------------------ | --- | --- | --- | ------------ | --- | --- | --- | --- |
• AssemblyDatumSide—Orientationof theassembly.This canhavethe
followingvalues:
|     | ○ Yellow—The                    |           | primarysideof |           | thedatumplanewhichis |     |     | thedefault |     |
| --- | ------------------------------- | --------- | ------------- | --------- | -------------------- | --- | --- | ---------- | --- |
|     | directionof                     | thearrow. |               |           |                      |     |     |            |     |
|     | ○ Red—Thesecondary              |           |               | sideofthe | datumplanewhichisthe |     |     | direction  |     |
|     | oppositeto                      | thatofthe |               | arrow.    |                      |     |     |            |     |
| •   | ComponentReference—Areferenceon |           |               |           | theplacedcomponent.  |     |     |            |     |
• ComponentDatumSide—Orientationof theassemblycomponent. This can
havethefollowingvalues:
|     | ○ Yellow—The                       |           | primarysideof        |           | thedatumplanewhichis |               |     | thedefault |     |
| --- | ---------------------------------- | --------- | -------------------- | --------- | -------------------- | ------------- | --- | ---------- | --- |
|     | directionof                        | thearrow. |                      |           |                      |               |     |            |     |
|     | ○ Red—Thesecondary                 |           |                      | sideofthe | datumplanewhichisthe |               |     | direction  |     |
|     | oppositeto                         | thatofthe |                      | arrow.    |                      |               |     |            |     |
| •   | Offset—Themateor                   |           | alignoffsetvaluefrom |           |                      | thereference. |     |            |     |
| •   | Attributes—Constraintattributesfor |           |                      |           | agivenconstraint     |               |     |            |     |
• UserDefinedData—Astringthatspecifiesuserdatafor thegiven constraint.
Usethe propertieslisted aboveto accessthe parametersof thecomponent
constraintobject.
| 258 |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- | --- | --- |

| Redefining | and | Rerouting | Assembly |     |
| ---------- | --- | --------- | -------- | --- |
Components
Thesefunctionsenableyouto reroutepreviouslyassembledcomponents,just asin
aninteractiveCreoParametricsession.
MethodsIntroduced:
• IpfcComponentFeat.RedefineThroughUI()
• IpfcComponentFeat.MoveThroughUI()
| ThemethodIpfcComponentFeat.RedefineThroughUI() |     |     |     | mustbe used |
| ---------------------------------------------- | --- | --- | --- | ----------- |
ininteractiveVBAPI applications.This methoddisplaysthe CreoParametric
Constraintdialogbox.This enablestheend usertoredefine theconstraints
interactively. Thecontrolreturns tothe VBAPI applicationwhen theuserselects
| OKorCancelandthe                           | dialogboxisclosed. |     |         |         |
| ------------------------------------------ | ------------------ | --- | ------- | ------- |
| ThemethodIpfcComponentFeat.MoveThroughUI() |                    |     | invokes | adialog |
boxthatpromptsthe userto interactivelyrepositionthe components.This
interfaceenablestheuserto specifythe translationandrotation values.The
controlreturnsto theVB APIapplicationwhenthe userselectsOK orCanceland
thedialogboxisclosed.
| Example:        | Component                              | Constraints |     |     |
| --------------- | -------------------------------------- | ----------- | --- | --- |
| Thesamplecodein | thefilepfcAssembliesExamples.vblocated |             |     | at  |
<creo_vbapi_loadpoint>/vbapi_examplesdisplayseachconstraint of
thecomponentvisuallyon thescreen,andincludes atextexplanationfor each
constraint.
| Example:                                          | Assembling                             | Components |     |       |
| ------------------------------------------------- | -------------------------------------- | ---------- | --- | ----- |
| Thesamplecodein                                   | thefilepfcAssembliesExamples.vblocated |            |     | at    |
| <creo_vbapi_loadpoint>/vbapi_examplesdemonstrates |                                        |            |     | howto |
assembleacomponentintoan assembly, andhowto constrainthe componentby
aligningdatumplanes.Ifthe completesetof datumplanesisnot found,the
functionwillshowthecomponentconstraintdialogto theuserto allowthemto
adjusttheplacement.
| Exploded | Assemblies |     |     |     |
| -------- | ---------- | --- | --- | --- |
Thesemethodsenableyouto determineandchangetheexplodestatusof the
assemblyobject.
MethodsandPropertiesIntroduced:
• IpfcAssembly.IsExploded
| AssembliesandComponents |     |     |     | 259 |
| ----------------------- | --- | --- | --- | --- |

• IpfcAssembly.Explode()
• IpfcAssembly.UnExplode()
• IpfcAssembly.GetActiveExplodedState()
• IpfcAssembly.GetDefaultExplodedState()
• IpfcExplodedState.Activate()
ThemethodsIpfcAssembly.Explode() and
IpfcAssembly.UnExplode()enableyouto determineandchangethe
explodestatusofthe assemblyobject.
ThemethodIpfcAssembly.GetActiveExplodedState() reports
whetherthe specifiedassemblyiscurrentlyexploded. Usethis methodin the
assemblymodeonly.The exploded statusofan assemblydepends on themode.If
anassembly isopened in thedrawing mode,thestateofthe assemblyin the
drawingview isdisplayed. Thedrawing view doesnotrepresent theactual
explodedstateof theassembly.
ThemethodIpfcAssembly.GetActiveExplodedState() returnsthe
currentactiveexplode state.
ThemethodIpfcAssembly.GetDefaultExplodedState() returnsthe
defaultexplodestate.
ThemethodIpfcExplodedState.Activate() activates thespecified
explodestaterepresentation.
Skeleton Models
Skeletonmodelsarea3-dimensional layoutof theassembly.These modelsare
holdersordistributors ofcriticaldesign information,andcanrepresent space
requirements,importantmounting locations,and motion.
MethodsandPropertiesIntroduced:
• IpfcAssembly.AssembleSkeleton()
• IpfcAssembly.AssembleSkeletonByCopy()
• IpfcAssembly.GetSkeleton()
• IpfcAssembly.DeleteSkeleton()
• IpfcSolid.IsSkeleton
ThemethodIpfcAssembly.AssembleSkeleton() addsan existing
skeletonmodelto thespecifiedassembly.
ThemethodIpfcAssembly.GetSkeleton() returnsthe skeletonmodelof
thespecifiedassembly.
ThemethodIpfcAssembly.DeleteSkeleton() deletesaskeleton model
componentfrom thespecifiedassembly.
260 Creo®ParametricVBAPIUser’sGuide

| ThemethodIpfcAssembly.AssembleSkeletonByCopy() |     | addsa |     |
| ---------------------------------------------- | --- | ----- | --- |
specifiedskeleton modeltothe assembly. Theinput parametersfor thismethod
are:
• SkeletonToCopy—Specifytheskeleton modeltobe copiedintotheassembly
| • NewSkeletonName—Specifyanamefor | thecopiedskeleton | model |     |
| --------------------------------- | ----------------- | ----- | --- |
ThepropertyIpfcSolid.IsSkeleton
|     | determinesifthe | specifiedpart | model |
| --- | --------------- | ------------- | ----- |
isaskeleton modelor aconcept model.It returnsatrue ifthe modelis askeleton
elseitreturnsa false.
| AssembliesandComponents |     |     | 261 |
| ----------------------- | --- | --- | --- |

21
Family Tables
WorkingwithFamilyTables......................................................................................263
CreatingFamilyTableInstances...............................................................................265
CreatingFamilyTableColumns................................................................................265
Thischapterdescribes howtousetheVB API classesandmethodsto accessand
manipulatefamilytable information.
262 Creo®ParametricVBAPIUser’sGuide

| Working | with | Family | Tables |     |     |
| ------- | ---- | ------ | ------ | --- | --- |
TheVB APIprovidesseveralmethodsfor accessing familytableinformation.
Becauseeverymodelinheritsfromthe interfaceIpfcFamilyMember,every
| modelcan  | havea familytableassociated |     | withit. |     |     |
| --------- | --------------------------- | --- | ------- | --- | --- |
| Accessing | Instances                   |     |         |     |     |
MethodsandPropertiesIntroduced:
• IpfcFamilyMember.Parent
• IpfcFamilyMember.GetImmediateGenericInfo()
• IpfcFamilyMember.GetTopGenericInfo()
• IpfcFamilyTableRow.CreateInstance()
• IpfcFamilyMember.ListRows()
• IpfcFamilyMember.GetRow()
• IpfcFamilyMember.RemoveRow()
• IpfcFamilyTableRow.InstanceName
• IpfcFamilyTableRow.IsLocked
| Toget | thegenericmodelfor | aninstance, | calltheproperty |     |     |
| ----- | ------------------ | ----------- | --------------- | --- | --- |
IpfcFamilyMember.Parent.
FromPro/ENGINEERWildfire4.0onwards,the behaviorofthe property
| IpfcFamilyMember.Parenthaschangedasa |     |     |     | resultofperformance |     |
| ------------------------------------ | --- | --- | --- | ------------------- | --- |
improvementinfamily tableretrievalmechanism.Whenyounowcallthe
propertyIpfcFamilyMember.Parent,
|     |     |     | itthrows | an exception |     |
| --- | --- | --- | -------- | ------------ | --- |
IpfcXToolkitCantOpen,ifthe immediategeneric ofamodelinstance in a
nestedfamilytable iscurrentlynotin session.Handle this exceptionandusethe
methodIpfcFamilyMember.GetImmediateGenericInfo() to getthe
modeldescriptorof theimmediategenericmodel.This informationcan beused to
retrievetheimmediategenericmodel.
Ifyouwishto switchoffthe abovebehaviorandcontinueto runlegacy
| applicationsin                 | thepre-Wildfire4.0mode,setthe |     |     | configurationoption     |     |
| ------------------------------ | ----------------------------- | --- | --- | ----------------------- | --- |
| retrieve_instance_dependencies |                               |     |     | toinstance_and_generic_ |     |
deps.
| Toget | themodeldescriptorofthetop |     | genericmodel,callthe |     | method |
| ----- | -------------------------- | --- | -------------------- | --- | ------ |
IpfcFamilyMember.GetTopGenericInfo().
Similarly, themethodIpfcFamilyTableRow.CreateInstance() returns
| aninstance | modelcreatedfrom | theinformationstoredin |     | the |     |
| ---------- | ---------------- | ---------------------- | --- | --- | --- |
IpfcFamilyTableRowobject.
FamilyTables 263

| ThemethodIpfcFamilyMember.ListRows() |     |     | returnsasequence | ofall |
| ------------------------------------ | --- | --- | ---------------- | ----- |
rowsin thefamilytable, whereasIpfcFamilyMember.GetRow() getsthe
| rowobjectwith | thenameyouspecify. |     |     |     |
| ------------- | ------------------ | --- | --- | --- |
Usethe methodIpfcFamilyMember.RemoveRow() topermanentlydelete
| therow fromthefamily |     | table. |     |     |
| -------------------- | --- | ------ | --- | --- |
ThepropertyIpfcFamilyTableRow.InstanceName returnsthename that
| correspondsto | theinvoking | rowobject. |     |     |
| ------------- | ----------- | ---------- | --- | --- |
Tocontrolwhetherthe instancecan bechanged orremoved,callthe property
IpfcFamilyTableRow.IsLocked.
| Accessing | Columns |     |     |     |
| --------- | ------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcFamilyMember.ListColumns()
• IpfcFamilyMember.GetColumn()
• IpfcFamilyMember.RemoveColumn
• IpfcFamilyTableColumn.Symbol
• IpfcFamilyTableColumn.Type
• IpfcFamColModelItem.RefItem
• IpfcFamColParam.RefParam
ThemethodIpfcFamilyMember.ListColumns()
returns asequenceof all
| columnsin                             | thefamilytable.  |     |         |              |
| ------------------------------------- | ---------------- | --- | ------- | ------------ |
| ThemethodIpfcFamilyMember.GetColumn() |                  |     | returns | afamilytable |
| column,given                          | itssymbolicname. |     |         |              |
Topermanentlydeletethecolumn fromthe familytableand allchangedvalues in
| allinstances,callthe | methodIpfcFamilyMember.RemoveColumn. |     |     |     |
| -------------------- | ------------------------------------ | --- | --- | --- |
ThepropertyIpfcFamilyTableColumn.Symbol returnsthe stringsymbol
|     | column,suchasD4 | orF5. |     |     |
| --- | --------------- | ----- | --- | --- |
atthe topofthe
ThepropertyIpfcFamilyTableColumn.Type returnsan enumeratedvalue
indicatingthetypeofparametergoverned bythe columnin thefamilytable.
| ThepropertyIpfcFamColModelItem.RefItem |     |              | returnsthe   |            |
| -------------------------------------- | --- | ------------ | ------------ | ---------- |
| IpfcModelItem(Feature                  |     | orDimension) | controlledby | thecolumn, |
whereasIpfcFamColParam.RefParam returnstheParameter controlled
bythecolumn.
| Accessing | Cell | Information |     |     |
| --------- | ---- | ----------- | --- | --- |
MethodsandPropertiesIntroduced:
| 264 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

• IpfcFamilyMember.GetCell()
• IpfcFamilyMember.GetCellIsDefault()
• IpfcFamilyMember.SetCell()
• IpfcParamValue.StringValue
• IpfcParamValue.IntValue
• IpfcParamValue.DoubleValue
• IpfcParamValue.BoolValue
ThemethodIpfcFamilyMember.GetCell()
returnsa string
IpfcParamValuethatcorrespondsto the cellattheintersectionof therow and
| columnarguments.Use                   | themethod |     |     |                  |       |
| ------------------------------------- | --------- | --- | --- | ---------------- | ----- |
| IpfcFamilyMember.GetCellIsDefault()to |           |     |     | checkif thevalue | ofthe |
specifiedcellis thedefaultvalue,whichis thevalue ofthespecifiedcellin the
genericmodel.
ThemethodIpfcFamilyMember.SetCell()
|                                |                |     |                          | assignsavalue to | acolumnin |
| ------------------------------ | -------------- | --- | ------------------------ | ---------------- | --------- |
| aparticularfamily              | tableinstance. |     |                          |                  |           |
| TheIpfcParamValue.StringValue, |                |     | IpfcParamValue.IntValue, |                  |           |
IpfcParamValue.DoubleValue,andIpfcParamValue.BoolValue
| propertiesare | usedto getthedifferent | typesofparameter |           | values. |     |
| ------------- | ---------------------- | ---------------- | --------- | ------- | --- |
| Creating      | Family                 | Table            | Instances |         |     |
MethodsIntroduced:
• IpfcFamilyMember.AddRow()
• CMpfcModelItem.CreateStringParamValue()
• CMpfcModelItem.CreateIntParamValue()
• CMpfcModelItem.CreateDoubleParamValue()
• CMpfcModelItem.CreateBoolParamValue()
| Usethe methodIpfcFamilyMember.AddRow()to |     |     |     | createanewinstance |     |
| ---------------------------------------- | --- | --- | --- | ------------------ | --- |
withthespecifiedname,and,optionally, thespecifiedvaluesfor eachcolumn.If
youdonotpassin asetof values,thevalue* willbeassignedto eachcolumn.
| Thisvalueindicates | thattheinstance | usesthe | genericvalue. |     |     |
| ------------------ | --------------- | ------- | ------------- | --- | --- |
| Creating           | Family          | Table   | Columns       |     |     |
MethodsIntroduced:
• IpfcFamilyMember.CreateDimensionColumn()
• IpfcFamilyMember.CreateParamColumn()
| FamilyTables |     |     |     |     | 265 |
| ------------ | --- | --- | --- | --- | --- |

• IpfcFamilyMember.CreateFeatureColumn()
• IpfcFamilyMember.CreateComponentColumn()
• IpfcFamilyMember.CreateCompModelColumn()
• IpfcFamilyMember.CreateGroupColumn()
• IpfcFamilyMember.CreateMergePartColumn()
• IpfcFamilyMember.CreateColumn()
• IpfcFamilyMember.AddColumn()
• IpfcpfcModelItem.CreateStringParamValue()
Theabovemethods initializeacolumn basedon theinput argument.These
| methodsassignthepropersymbol             |     |     | to thecolumnheader. |                   |     |
| ---------------------------------------- | --- | --- | ------------------- | ----------------- | --- |
| ThemethodIpfcFamilyMember.CreateColumn() |     |     |                     | createsanewcolumn |     |
givenaproperlydefinedsymbol andcolumntype. Theresults ofthiscall should
| bepassedto                            | themethodIpfcFamilyMember.AddColumn() |     |     |                  | to addthe |
| ------------------------------------- | ------------------------------------- | --- | --- | ---------------- | --------- |
| columntothe                           | model'sfamilytable.                   |     |     |                  |           |
| ThemethodIpfcFamilyMember.AddColumn() |                                       |     |     | adds thecolumnto | the       |
familytable.You canspecify thevalues;if youpassnothingfor thevalues,the
methodassignsthevalue* toeachinstance to acceptthecolumn’s defaultvalue.
| Example | Code: | Adding | Dimensions | to a | Family |
| ------- | ----- | ------ | ---------- | ---- | ------ |
Table
| Thesamplecodein | thefilepfcFamilyTablesExamples.vb |     |     |     | locatedat |
| --------------- | --------------------------------- | --- | --- | --- | --------- |
<creo_vbapi_loadpoint>/vbapi_examplesadds allthedimensions to
afamilytable.The programlists thedependenciesof theassembly andloops
througheachdependency,assigning themodelto anewFamColDimension
columnobject.All thedimensions,parameters,features,andcomponentscouldbe
| addedtothe | familytable | usinga | similarmethod. |                                 |     |
| ---------- | ----------- | ------ | -------------- | ------------------------------- | --- |
| 266        |             |        |                | Creo®ParametricVBAPIUser’sGuide |     |

22
Action Listeners
TheVBAPIActionListeners....................................................................................268
ActionSources........................................................................................................268
TypesofActionListeners.........................................................................................269
CancellinganActionListenerOperation.....................................................................275
Thischapterdescribes theVB APImethodsthatenableyouto useactionlisteners.
267

| The VB | API | Action |     | Listeners |     |     |     |     |
| ------ | --- | ------ | --- | --------- | --- | --- | --- | --- |
AnActionListeneris aclass thatisassignedto respondto certainevents.In
theVB API,youcanassignactionlistenersto respond toevents involvingthe
followingtasks:
• Changingwindows
• Changingworkingdirectories
• Modeloperations
• Regenerating
| • Creating,deleting,andredefining |     |       |               |     | features |                 |     |     |
| --------------------------------- | --- | ----- | ------------- | --- | -------- | --------------- | --- | --- |
| • Checkingforregeneration         |     |       | failures      |     |          |                 |     |     |
| Allactionlistenersin              |     | theVB | APIaredefined |     |          | bytheseclasses: |     |     |
• Interface—Named<Object>ActionListener. Thisinterfacedefines the
| methodsthatcanrespond |     |     | tovarious |     | events. |     |     |     |
| --------------------- | --- | --- | --------- | --- | ------- | --- | --- | --- |
• Defaultclass—NamedDefault<Object>ActionListener. Thisclass hasevery
availablemethod overriddenbyan emptyimplementation.Youcreateyour
ownactionlistenersby extendingthedefaultclass andoverridingthe methods
| forevents | thatinterestyou. |     |     |     |     |     |     |     |
| --------- | ---------------- | --- | --- | --- | --- | --- | --- | --- |
Note
Whennotificationsaresetin VB APIapplications, everytimean eventis
| triggered, | notificationmessages |     |     | areaddedto |     | thetrailfiles. |     | From Creo |
| ---------- | -------------------- | --- | --- | ---------- | --- | -------------- | --- | --------- |
variablePROTK_LOG_
| Parametric2.0M210                |                         | onward,anewenvironment |     |     |                    |             |     |           |
| -------------------------------- | ----------------------- | ---------------------- | --- | --- | ------------------ | ----------- | --- | --------- |
| DISABLE                          | enablesyoutodisablethis |                        |     |     | behavior.Whensetto |             |     | true, the |
| notificationsmessagesarenotadded |                         |                        |     |     | tothe              | trailfiles. |     |           |
| Action                           | Sources                 |                        |     |     |                    |             |     |           |
Methodsintroduced:
• IpfcActionSource.AddActionListener
• IpfcActionSource.AddActionListenerWithType
• IpfcActionSource.RemoveActionListener
ManyVB APIclasses inherittheIpfcActionSource interface,but onlythe
| followingclasses |     | currentlymake |     | callsto | themethodsof |     | registered |     |
| ---------------- | --- | ------------- | --- | ------- | ------------ | --- | ---------- | --- |
IpfcActionListeners:
• IpfcSession
| 268 |     |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

| ○ SessionAction |          | Listener |     |     |
| --------------- | -------- | -------- | --- | --- |
| ○ ModelAction   |          | Listener |     |     |
| ○ SolidAction   | Listener |          |     |     |
| ○ ModelEvent    | Action   | Listener |     |     |
| ○ FeatureAction |          | Listener |     |     |
• IpfcUICommand
| ○ UIAction      | Listener |                  |     |     |
| --------------- | -------- | ---------------- | --- | --- |
| • IpfcModel(and |          | it’s subclasses) |     |     |
| ○ ModelAction   |          | Listener         |     |     |
○
| ParameterAction |          | Listener         |     |     |
| --------------- | -------- | ---------------- | --- | --- |
| • IpfcSolid(and |          | it’s subclasses) |     |     |
| ○ SolidAction   | Listener |                  |     |     |
| ○ FeatureAction |          | Listener         |     |     |
| • IpfcFeature   | (andit’s | subclasses)      |     |     |
| ○ FeatureAction |          | Listener         |     |     |
Note
○ Assigningan actionlistenerto asourcenot relatedto itwillnotcausean
| errorbut          | thelistenermethod                        |        | willneverbe | called. |
| ----------------- | ---------------------------------------- | ------ | ----------- | ------- |
| ○ The             | methodIpfcActionSource.AddActionListener |        |             | adds    |
| actionlistenersto |                                          | notify | youofallthe | events. |
○
| The | method |     |     |     |
| --- | ------ | --- | --- | --- |
IpfcActionSource.AddActionListenerWithType addsaction
listenersto notify youabouteventsthatarespecifiedin theinput argument
ATypes.
| Types | of Action |     | Listeners |     |
| ----- | --------- | --- | --------- | --- |
Thefollowingsectionsdescribethe differentkinds ofactionlisteners:session, UI
| command,solid, | andfeature. |        |           |     |
| -------------- | ----------- | ------ | --------- | --- |
| Session        | Level       | Action | Listeners |     |
Methodsintroduced:
ActionListeners 269

• IpfcSessionActionListener.OnAfterDirectoryChange()
• IpfcSessionActionListener.OnAfterWindowChange()
• IpfcSessionActionListener.OnAfterModelDisplay()
• IpfcSessionActionListener.OnBeforeModelErase()
• IpfcSessionActionListener.OnBeforeModelDelete()
• IpfcSessionActionListener.OnBeforeModelRename()
• IpfcSessionActionListener.OnBeforeModelSave()
• IpfcSessionActionListener.OnBeforeModelPurge()
• IpfcSessionActionListener.OnBeforeModelCopy()
• IpfcSessionActionListener.OnAfterModelPurge()
TheIpfcSessionActionListener.OnAfterDirectoryChange()
methodactivatesaftertheuserchangesthe workingdirectory. Thismethod takes
thenewdirectorypathasan argument.
TheIpfcSessionActionListener.OnAfterWindowChange()
methodactivateswhen theuseractivatesa windowotherthanthecurrentone.
Passthe newwindowto themethodasan argument.
TheIpfcSessionActionListener.OnAfterModelDisplay()
methodactivateseverytime amodelisdisplayedin awindow.
Note
Modeldisplayevents happenwhen windowsaremoved, openedandclosed,
repainted,orthe modelis regenerated.Theevent canoccurmorethanoncein
succession.
Themethods
IpfcSessionActionListener.OnBeforeModelErase(),
IpfcSessionActionListener.OnBeforeModelRename,
IpfcSessionActionListener.OnBeforeModelSave,and
IpfcSessionActionListener.OnBeforeModelCopy takespecial
arguments.Theyaredesigned toallow youtofillin theargumentsandpassthis
databackto Creo Parametric.The modelnamesplacedin thedescriptorswillbe
usedbyCreo Parametricasthe defaultnamesin theuserinterface.
UI Command Action Listeners
Methodsintroduced:
• IpfcSession.UICreateCommand()
270 Creo®ParametricVBAPIUser’sGuide

• IpfcUICommandActionListener.OnCommand()
TheIpfcSession.UICreateCommand()
methodtakesa
IpfcUICommandActionListenerargumentandreturnsa
IpfcUICommandactionsourcewith thatactionlisteneralreadyregistered.This
| IpfcUICommandobjectis                    |     | subsequentlypassedas | an argumentto | the |
| ---------------------------------------- | --- | -------------------- | ------------- | --- |
| IpfcSession.UIAddButton()methodthataddsa |     |                      | commandbutton | toa |
CreoParametricmenu.The
| IpfcUICommandActionListener.OnCommandmethod |     |     |                  | oftheregistered |
| ------------------------------------------- | --- | --- | ---------------- | --------------- |
| IpfcUICommandActionListeneriscalledwhenever |     |     | thecommandbutton |                 |
isclicked.
| Model | Level Action | listeners |     |     |
| ----- | ------------ | --------- | --- | --- |
Methodsintroduced:
• IpfcModelActionListener.OnAfterModelSave()
• IpfcModelEventActionListener.OnAfterModelCopy()
• IpfcModelEventActionListener.OnAfterModelRename()
• IpfcModelEventActionListener.OnAfterModelErase()
• IpfcModelEventActionListener.OnAfterModelDelete()
• IpfcModelActionListener.OnAfterModelRetrieve()
• IpfcModelActionListener.OnBeforeModelDisplay()
• IpfcModelActionListener.OnAfterModelCreate()
• IpfcModelActionListener.OnAfterModelSaveAll()
• IpfcModelEventActionListener.OnAfterModelCopyAll()
• IpfcModelActionListener.OnAfterModelEraseAll()
• IpfcModelActionListener.OnAfterModelDeleteAll()
• IpfcModelActionListener.OnAfterModelRetrieveAll()
Methodsending in Allarecalledafteranyevent ofthe specifiedtype. Thecall is
madeevenifthe userdidnotexplicitlyrequest thattheactiontakeplace.Methods
thatdo notendin Allare onlycalledwhen theuserspecificallyrequests thatthe
eventoccurs.
ThemethodIpfcModelActionListener.OnAfterModelSave()
is
| calledaftersuccessfullysavinga |     | model. |     |     |
| ------------------------------ | --- | ------ | --- | --- |
Themethod
| IpfcModelEventActionListener.OnAfterModelCopy() |                |     |     | iscalled |
| ----------------------------------------------- | -------------- | --- | --- | -------- |
| aftersuccessfully                               | copyingamodel. |     |     |          |
ActionListeners 271

Themethod
| IpfcModelEventActionListener.OnAfterModelRename() |                 |     | iscalled |     |
| ------------------------------------------------- | --------------- | --- | -------- | --- |
| aftersuccessfully                                 | renamingamodel. |     |          |     |
Themethod
| IpfcModelEventActionListener.OnAfterModelErase()is |         |         |     | called |
| -------------------------------------------------- | ------- | ------- | --- | ------ |
| aftersuccessfully                                  | erasing | amodel. |     |        |
Themethod
IpfcModelEventActionListener.OnAfterModelDelete()
iscalled
| aftersuccessfully | deletingamodel. |     |     |     |
| ----------------- | --------------- | --- | --- | --- |
ThemethodIpfcModelActionListener.OnAfterModelRetrieve()
| iscalledaftersuccessfullyretrieving |     | amodel. |     |     |
| ----------------------------------- | --- | ------- | --- | --- |
ThemethodIpfcModelActionListener.OnBeforeModelDisplay()
| iscalledbeforedisplaying |     | amodel. |     |     |
| ------------------------ | --- | ------- | --- | --- |
Note
The method
| IpfcModelActionListener.OnBeforeModelDisplay()        |                   |     | isnot |     |
| ----------------------------------------------------- | ----------------- | --- | ----- | --- |
| supportedin                                           | asynchronousmode. |     |       |     |
| ThemethodIpfcModelActionListener.OnAfterModelCreate() |                   |     |       | is  |
calledafterthesuccessfulcreationofa model.
| Solid Level | Action | Listeners |     |     |
| ----------- | ------ | --------- | --- | --- |
Methodsintroduced:
• IpfcSolidActionListener.OnBeforeRegen()
• IpfcSolidActionListener.OnAfterRegen()
• IpfcSolidActionListener.OnBeforeUnitConvert()
• IpfcSolidActionListener.OnAfterUnitConvert()
• IpfcSolidActionListener.OnBeforeFeatureCreate()
• IpfcSolidActionListener.OnAfterFeatureCreate()
• IpfcSolidActionListener.OnAfterFeatureDelete()
| TheIpfcSolidActionListener.OnBeforeRegen() |     |     | and              |     |
| ------------------------------------------ | --- | --- | ---------------- | --- |
| IpfcSolidActionListener.OnAfterRegen()     |     |     | methodsoccurwhen | the |
userregeneratesa solidobjectwithin theIpfcActionSource to whichthe
listenerisassigned. Thesemethodstakethe firstfeatureto be regeneratedanda
| 272 |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | ------------------------------- | --- | --- |

handleto theIpfcSolid objectasarguments. Inaddition, themethod
| IpfcSolidActionListener.OnAfterRegenerate()        |     |     |     |     | includesa   |
| -------------------------------------------------- | --- | --- | --- | --- | ----------- |
| Booleanargumentthatindicateswhetherregenerationwas |     |     |     |     | successful. |
Note
| • Itis notrecommendedto |     | modifygeometryordimensionsusing |     |     | the |
| ----------------------- | --- | ------------------------------- | --- | --- | --- |
IpfcSolidActionListener.OnBeforeRegenerate()method call.
• Aregenerationthatdidnot takeplacebecausenothingwas modifiedis
| identifiedasaregeneration |     | failure. |     |     |     |
| ------------------------- | --- | -------- | --- | --- | --- |
ThemethodsIpfcSolidActionListener.OnBeforeUnitConvert()
andIpfcSolidActionListener.OnAfterUnitConvert()activate
whenausermodifies theunitscheme(byselectingthe CreoParametriccommand
theSolid
| SetUp, Units). | Themethodsreceive |     |     | objecttobe | convertedanda |
| -------------- | ----------------- | --- | --- | ---------- | ------------- |
Booleanflag thatidentifies whethertheconversionchangedthe dimensionvalues
| tokeeptheobjectthe |     | samesize. |     |     |     |
| ------------------ | --- | --------- | --- | --- | --- |
Note
IpfcSolidActionListeners canbe registeredwith thesessionobject
sothatitsmethodsare calledwhentheseevents occurforanysolidmodelthat
| isin session. |     |     |     |     |     |
| ------------- | --- | --- | --- | --- | --- |
TheIpfcSolidActionListener.OnBeforeFeatureCreate()
methodactivateswhen theuserstartsto createa featurethatrequirestheFeature
Creationdialogbox. Becausethisevent occursonlyafterthe dialogboxis
displayed,itwillnotoccuratallfor datumsandotherfeaturesthatdo notusethis
dialogbox. This methodtakestwoarguments:the solidmodelthatwillcontain
| thefeatureandtheIpfcModelItem |     |     | identifier. |     |     |
| ----------------------------- | --- | --- | ----------- | --- | --- |
TheIpfcSolidActionListener.OnAfterFeatureCreate() method
activatesafteranyfeature,includingdatums, hasbeencreated.This methodtakes
| thenewIpfcFeature |     | objectasan | argument. |     |     |
| ----------------- | --- | ---------- | --------- | --- | --- |
TheIpfcSolidActionListener.OnAfterFeatureDelete() method
activatesafteranyfeaturehasbeendeleted.Themethod receivesthesolidthat
containedthefeatureandthe (nowdefunct)IpfcModelItem identifier.
| Selection | Level | Action | Listeners |     |     |
| --------- | ----- | ------ | --------- | --- | --- |
Methodintroduced:
ActionListeners 273

• IpfcSelectionBufferListener.OnAfterSelBufferChange()
ThemethodIpfcSelectionBufferListener.OnAfterSelBufferChange() iscalled
| afterthe | selectionbuffer ischanged. |           |     |
| -------- | -------------------------- | --------- | --- |
| Feature  | Level Action               | Listeners |     |
Methodsintroduced:
• IpfcFeatureActionListener.OnBeforeDelete()
• IpfcFeatureActionListener.OnBeforeSuppress()
• IpfcFeatureActionListener.OnAfterSuppress()
• IpfcFeatureActionListener.OnBeforeRegen()
• IpfcFeatureActionListener.OnAfterRegen()
• IpfcFeatureActionListener.OnRegenFailure()
• IpfcFeatureActionListener.OnBeforeRedefine()
• IpfcFeatureActionListener.OnAfterCopy()
• IpfcFeatureActionListener.OnBeforeParameterDelete()
| Eachmethodin            | IpfcFeatureActionListenertakesasan |     | argumentthe |
| ----------------------- | ---------------------------------- | --- | ----------- |
| featurethattriggeredthe | event.                             |     |             |
IpfcFeatureActionListenerscan beregistered withthe objectso that
theactionlistener’s methodsare calledwhenevertheseeventsoccurfor any
featurethatis in sessionorwith asolidmodelto react tochangesonly in that
model.
ThemethodIpfcFeatureActionListener.OnBeforeDelete() is
| calledbeforeafeatureis | deleted. |     |     |
| ---------------------- | -------- | --- | --- |
ThemethodIpfcFeatureActionListener.OnBeforeSuppress() is
| calledbeforeafeatureis | suppressed. |     |     |
| ---------------------- | ----------- | --- | --- |
ThemethodIpfcFeatureActionListener.OnAfterSuppressv is
calledafterasuccessfulfeaturesuppression.
ThemethodIpfcFeatureActionListener.OnBeforeRegen() is
| calledbeforeafeatureis                           | regenerated. |     |           |
| ------------------------------------------------ | ------------ | --- | --------- |
| ThemethodIpfcFeatureActionListener.OnAfterRegenv |              |     | is called |
aftera successfulfeatureregeneration.
ThemethodIpfcFeatureActionListener.OnRegenFailure() is
calledwhenafeaturefailsregeneration.
ThemethodIpfcFeatureActionListener.OnBeforeRedefine() is
| calledbeforeafeatureis | redefined. |     |     |
| ---------------------- | ---------- | --- | --- |
274 Creo®ParametricVBAPIUser’sGuide

ThemethodIpfcFeatureActionListener.OnAfterCopy() is called
aftera featurehasbeensuccessfully copied.
Themethod
IpfcFeatureActionListener.OnBeforeParameterDelete() is
calledbeforeafeatureparameter isdeleted.
Cancelling an ActionListener Operation
TheVB APIallowsyou tocancelcertain notificationevents,registeredbythe
actionlisteners.
MethodsIntroduced:
• CCpfcXCancelProEAction.Throw()
ThestaticmethodCCpfcXCancelProEAction.Throw() mustbe called
fromthebody ofan actionlistenerto cancelthe impendingCreo Parametric
operation.This methodwillthrow atheVB APIexception signallingto Creo
Parametrictocancelthe listenerevent.
Note:Yourapplicationshouldnotcatch theVB APIexception,or shouldrethrow
itif caught,so thatCreoParametricis forcedto handleit.
Thefollowingeventscan becancelledusing thistechnique:
• IpfcSessionActionListener.OnBeforeModelErase()
• IpfcSessionActionListener.OnBeforeModelDelete()
• IpfcSessionActionListener.OnBeforeModelRename()
• IpfcSessionActionListener.OnBeforeModelSave()
• IpfcSessionActionListener.OnBeforeModelPurge()
• IpfcSessionActionListener.OnBeforeModelCopy()
• IpfcModelActionListener.OnBeforeParameterCreate()
• IpfcModelActionListener.OnBeforeParameterDelete()
• IpfcModelActionListener.OnBeforeParameterModify()
• IpfcFeatureActionListener.OnBeforeDelete()
• IpfcFeatureActionListener.OnBeforeSuppress()
• IpfcFeatureActionListener.OnBeforeParameterDelete()
• IpfcFeatureActionListener.OnBeforeParameterCreate()
• IpfcFeatureActionListener.OnBeforeRedefine()
ActionListeners 275

23
Interface
ExportingFilesand2DModels.................................................................................277
ExportingtoPDFandU3D.......................................................................................284
Exporting3DGeometry...........................................................................................291
ShrinkwrapExport...................................................................................................294
ImportingFiles........................................................................................................298
Importing3DGeometry............................................................................................301
PlottingFiles...........................................................................................................303
PrintingFiles...........................................................................................................304
SolidOperations......................................................................................................311
WindowOperations.................................................................................................312
Thischapterdescribes variousmethodsof importingand exportingfilesin theVB
API.
276 Creo®ParametricVBAPIUser’sGuide

| Exporting | Files | and 2D | Models |     |
| --------- | ----- | ------ | ------ | --- |
MethodIntroduced:
• IpfcModel.Export()
ThemethodIpfcModel.Export() exportsmodeldatato afile. Theexported
filesareplacedin thecurrentCreo Parametricworking directory. Theinput
parametersare:
| • filename—Outputfilenameincluding |     |     | extensions |     |
| ---------------------------------- | --- | --- | ---------- | --- |
• exportdata—TheIpfcExportInstructions objectthatcontrols the
| exportoperation.The | typeof | datathatisexportedis |     | givenby the |
| ------------------- | ------ | -------------------- | --- | ----------- |
IpfcExportTypeobject.
Therearefourgeneralcategories offilesto whichyoucan exportmodels:
| • Filetypeswhose | instructionsinheritfrom |     |     |     |
| ---------------- | ----------------------- | --- | --- | --- |
IpfcGeomExportInstructions.
Theseinstructionsexportfilesthatcontainprecisegeometricinformationused
byotherCADsystems.
| • Filetypeswhose | instructionsinheritfrom |     |     |     |
| ---------------- | ----------------------- | --- | --- | --- |
IpfcCoordSysExportInstructions.
Theseinstructionsexportfilesthatcontaincoordinateinformationdescribing
faceted,solidmodels(withoutdatumsandsurfaces).
| • Filetypeswhose | instructionsinheritfrom |     |     |     |
| ---------------- | ----------------------- | --- | --- | --- |
IpfcFeatIdExportInstructions.
| Theseinstructionsexportinformationabout |     |     | aspecificfeature. |     |
| --------------------------------------- | --- | --- | ----------------- | --- |
• Generalfiletypesthatinherit onlyfrom IpfcExportInstructions.
Theseinstructionsprovideconversionstofiletypessuchas BOM(billof
materials).
Forinformationon exportingto aspecificformat, seetheVB APIAPIWizardand
| onlinehelpfor theCreo | Parametricinterface. |     |     |     |
| --------------------- | -------------------- | --- | --- | --- |
Export Instructions
MethodsIntroduced:
• CCpfcRelationExportInstructions.Create()
• CCpfcModelInfoExportInstructions.Create()
• CCpfcProgramExportInstructions.Create()
• CCpfcIGESFileExportInstructions.Create()
Interface 277

• CCpfcDXFExportInstructions.Create()
• CCpfcRenderExportInstructions.Create()
• CCpfcSTLASCIIExportInstructions.Create()
• CCpfcSTLBinaryExportInstructions.Create()
• CCpfcBOMExportInstructions.Create()
• CCpfcDWGSetupExportInstructions.Create()
• CCpfcFeatInfoExportInstructions.Create()
• CCpfcMFGFeatCLExportInstructions.Create()
• CCpfcMFGOperCLExportInstructions.Create()
• CCpfcMaterialExportInstructions.Create()
• CCpfcCGMFILEExportInstructions.Create()
• CCpfcInventorExportInstructions.Create()
• CCpfcFIATExportInstructions.Create()
• CCpfcConnectorParamExportInstructions.Create()
• CCpfcCableParamsFileInstructions.Create()
• CCpfcCATIAFacetsExportInstructions.Create()
• CCpfcVRMLModelExportInstructions.Create
• CCpfcSTEP2DExportInstructions.Create()
• CCpfcMedusaExportInstructions.Create()
• CCpfcCADDSExportInstructions.Create()
• CCpfcSliceExportData.Create()
• CCpfcNEUTRALFileExportInstructions.Create()
• CCpfcProductViewExportInstructions.Create()
• IpfcBaseSession.ExportDirectVRML()
Export Instructions Table
Interface UsedtoExport
IpfcRelationExportInstructions Alistoftherelationsandparametersinapartor
assembly
IpfcModelInfoExportInstructions Informationaboutamodel,includingunits
information,features,andchildren
IpfcProgramExportInstructions Aprogramfileforapartorassemblythatcanbe
editedtochangethemodel
IpfcIGESExportInstructions AdrawinginIGESformat
IpfcDXFExportInstructions AdrawinginDXFformat
IpfcRenderExportInstructions ApartorassemblyinRENDERformat
278 Creo®ParametricVBAPIUser’sGuide

Interface UsedtoExport
IpfcSTLASCIIExportInstructions ApartorassemblytoanASCIISTLfile
IpfcSTLBinaryExportInstructions ApartorassemblyinabinarySTLfile
IpfcBOMExportInstructions ABOMforanassembly
IpfcDWGSetupExportInstructions Adrawingsetupfile
IpfcFeatInfoExportInstructions Informationaboutonefeatureinapartorassembly
IpfcMfgFeatCLExportInstructions Acutterlocation(CL)fileforoneNCsequenceina
manufacturingassembly
IpfcMfgOperClExportInstructions Acutterlocation(CL)fileforalltheNCsequences
inamanufacturingassembly
IpfcMaterialExportInstructions Amaterialfromapart
IpfcCGMFILEExportInstructions AdrawinginCGMformat
IpfcInventorExportInstructions ApartorassemblyinInventorformat
IpfcFIATExportInstructions ApartorassemblyinFIATformat
IpfcConnectorParamExportInstruc Theparametersofaconnectortoatextfile
tions
IpfcCableParamsFileInstructions Cableparametersfromanassembly
IpfcCATIAFacetsExportInstructions ApartorassemblyinCATIAformat(asafaceted
model)
IpfcVRMLModelExportInstructions ApartorassemblyinVRMLformat
IpfcSTEP2DExportInstructions Atwo-dimensionalSTEPformatfile
IpfcMedusaExportInstructions AdrawinginMEDUSAfile
IpfcCADDSExportInstructions ACADDS5solidmodel
IpfcNEUTRALFileExportInstructions ACreoParametricparttoneutralformat
IpfcProductViewExportInstructions Apart,assembly,ordrawinginCreoViewformat
IpfcExport.SliceExportData Asliceexportformat
Note
The NewInstruction Classesreplacethe followingDeprecatedClasses:
DeprecatedClasses NewInstructionClasses
IpfcSTEPExportInstructions IpfcSTEP3DExportInstructions
IpfcVDAExportInstructions IpfcVDA3DExportInstructions
IpfcIGES3DExportInstructions IpfcIGES3DNewExportInstructions
Exporting Drawing Sheets
Theoptions requiredto exportmultiple sheetsofa drawingaregiven bythe
IpfcExport2DOptionobject.
MethodsandPropertiesIntroduced:
• CCpfcExport2DOption.Create()
Interface 279

• IpfcExport2DOption.ExportSheetOption
• IpfcExport2DOption.ModelSpaceSheet
• IpfcExport2DOption.Sheets
ThemethodCCpfcExport2DOptions.Create() createsanewinstance of
theIpfcExport2DOptionobject.This
objectcontainsthefollowingoptions:
• ExportSheetOption—Specifiestheoption forexportingmultiple drawing
| sheets.Usethe | propertyIpfcExport2DOption.ExportSheetOption |     |     |     |     |
| ------------- | -------------------------------------------- | --- | --- | --- | --- |
tosetthe optionfor exportingmultiple drawingsheets.The optionsaregiven
| bythe EpfcExport2DSheetOptionclass |     |     | andcanbe | ofthe following |     |
| ---------------------------------- | --- | --- | -------- | --------------- | --- |
types:
| ○ EpfcEXPORT_CURRENT_TO_MODEL_SPACE—Exports |     |     |     | onlythe |     |
| ------------------------------------------- | --- | --- | --- | ------- | --- |
drawing’s currentsheetasmodelspaceto asinglefile. This isthedefault
type.
○ EpfcEXPORT_CURRENT_TO_PAPER_SPACE—Exports
onlythe
| drawing’s                               | currentsheetaspaperspaceto |                            | asingle file.This | typeisthe       |     |
| --------------------------------------- | -------------------------- | -------------------------- | ----------------- | --------------- | --- |
| sameasEpfcEXPORT_CURRENT_TO_MODEL_SPACE |                            |                            |                   | for formatsthat |     |
| do notsupporttheconcept                 |                            | ofmodelspaceandpaperspace. |                   |                 |     |
○ EpfcEXPORT_ALL—Exportsall thesheetsin adrawing to asinglefile
| aspaperspace,if               |     | applicableforthe | formattype.      |          |     |
| ----------------------------- | --- | ---------------- | ---------------- | -------- | --- |
| ○ EpfcEXPORT_SELECTED—Exports |     |                  | selectedsheetsin | adrawing | as  |
| paperspaceandonesheetas       |     | modelspace.      |                  |          |     |
• ModelSpaceSheet—Specifiesthe sheetnumberthatneedsbe exported as
modelspace.This option isapplicableonly ifthe exportformatssupportthe
conceptof modelspaceand paperspaceandifExportSheetOptionis setto
| EpfcEXPORT_SELECTED.               |     | Usethe | property          |     |     |
| ---------------------------------- | --- | ------ | ----------------- | --- | --- |
| IpfcExport2DOption.ModelSpaceSheet |     |        | to setthisoption. |     |     |
• Sheets—Specifiesthesheetnumbersthatneedto be exportedaspaperspace.
Thisoption isapplicable onlyif ExportSheetOptionissetto EpfcEXPORT_
| SELECTED. | Usethe | propertyIpfcExport2DOption.Sheetsto |     |     | setthis |
| --------- | ------ | ----------------------------------- | --- | --- | ------- |
option.
| Exporting | to Faceted | Formats |     |     |     |
| --------- | ---------- | ------- | --- | --- | --- |
Themethodsand propertiesdescribed in thissectionsupport theexportof Creo
|                       |     |               | facetedformatslikeCATIA | CGR. |     |
| --------------------- | --- | ------------- | ----------------------- | ---- | --- |
| Parametricdrawingsand |     | solidmodelsto |                         |      |     |
MethodsandPropertiesIntroduced:
• IpfcTriangulationInstructions.AngleControl
• IpfcTriangulationInstructions.ChordHeight
280 Creo®ParametricVBAPIUser’sGuide

• IpfcTriangulationInstructions.StepSize
• IpfcTriangulationInstructions.FacetControlOptions
| ThepropertyIpfcTriangulationInstructions.AngleControl |     |     |     |     |     |     | allow |
| ----------------------------------------------------- | --- | --- | --- | --- | --- | --- | ----- |
youtocontrolthe anglecontrolfor theexportedfacet drawings andmodels.You
| cansetthevalue | between0.0to | 1.0. |     |     |     |     |     |
| -------------- | ------------ | ---- | --- | --- | --- | --- | --- |
Usethe propertyIpfcTriangulationInstructions.ChordHeight
allowyoutocontrol thechordheightfortheexported facetdrawings andmodels.
| ThepropertyIpfcTriangulationInstructions.StepSize |                 |                |     |                 |     |        | allowsyou |
| ------------------------------------------------- | --------------- | -------------- | --- | --------------- | --- | ------ | --------- |
| tocontrolthe                                      | stepsize forthe | exportedfiles. |     | Thedefaultvalue |     | is0.0. |           |
Note
Youmustpassthevalue ofStep Sizevalue asNULL, ifyouspecify the
Qualityvalue.
| ThepropertyIpfcCoordSysExportInstructions.StepSize |                       |     |                 |     |        |     | control |
| -------------------------------------------------- | --------------------- | --- | --------------- | --- | ------ | --- | ------- |
| thestepsize                                        | forthe exportedfiles. |     | Thedefaultvalue |     | is0.0. |     |         |
Note
Youmustpassthevalue ofStep Sizevalue asNULL, ifyouspecify the
Qualityvalue.
Theproperty
| IpfcTriangulationInstructions.FacetControlOptions |     |     |     |     |     |     | control |
| ------------------------------------------------- | --- | --- | --- | --- | --- | --- | ------- |
thefacetexport optionsusing bitflags. You cansetthebit flagsusing the
IpfcFacetControlFlagobject.Ithasthefollowingvalues:
• EpfcFACET_STEP_SIZE_ADJUST—Adjusts thestep sizeaccordingto the
componentsize.
| • EpfcFACET_CHORD_HEIGHT_ADJUST—Adjusts |                   |     |                             |        | thechordheight |          |     |
| --------------------------------------- | ----------------- | --- | --------------------------- | ------ | -------------- | -------- | --- |
| accordingto                             | thecomponentsize. |     |                             |        |                |          |     |
| • EpfcFACET_USE_CONFIG—If               |                   |     | thisflag                    | isset, | valuesof       | theflags |     |
| EpfcFACET_STEP_SIZE_OFF,                |                   |     | EpfcFACET_STEP_SIZE_ADJUST, |        |                |          |     |
| andEpfcFACET_CHORD_HEIGHT_ADJUSTare     |                   |     |                             |        | ignoredand     | the      |     |
configurationsettings fromthe CreoParametricuserinterfaceare usedduring
| theexport                                 | operation.                 |     |     |                  |                 |                   |       |
| ----------------------------------------- | -------------------------- | --- | --- | ---------------- | --------------- | ----------------- | ----- |
| • EpfcFACET_CHORD_HEIGHT_DEFAULT—Uses     |                            |     |     |                  | thedefaultvalue |                   | setin |
| theCreo                                   | Parametricuserinterfacefor |     |     | thechordheight.  |                 |                   |       |
| • EpfcFACET_ANGLE_CONTROL_DEFAULT—Usesthe |                            |     |     |                  |                 | defaultvaluesetin |       |
| theCreo                                   | Parametricuserinterfacefor |     |     | theanglecontrol. |                 |                   |       |
| Interface                                 |                            |     |     |                  |                 |                   | 281   |

| • EpfcFACET_STEP_SIZE_DEFAULT—Uses           |     |        | thedefault | valuesetin   | the |
| -------------------------------------------- | --- | ------ | ---------- | ------------ | --- |
| CreoParametricuserinterface                  |     | forthe | stepsize.  |              |     |
| • EpfcFACET_STEP_SIZE_OFF—Switchesoffthestep |     |        |            | sizecontrol. |     |
• EpfcFACET_FORCE_INTO_RANGE—Forces theout-of-range parameters
intorange.Ifany oftheEpfcFACET_*_DEFAULT option isset, thenthe
optionpfcFACET_FORCE_INTO_RANGE isnotapplied onthatparameter.
| • EpfcFACET_STEP_SIZE_FACET_INCLUDE_QUILTS—Includes |                       |                     |                     |               | quilts |
| --------------------------------------------------- | --------------------- | ------------------- | ------------------- | ------------- | ------ |
| inthe exportof                                      | CreoParametricmodelto |                     | thespecifiedformat. |               |        |
| • EpfcEXPORT_INCLUDE_ANNOTATIONS—Includes           |                       |                     |                     | annotationsin | the    |
| exportof                                            | CreoParametricmodelto | thespecifiedformat. |                     |               |        |
Note
Toincludeannotations, duringthe exportof CreoParametricmodel, you
| mustcallthemethod |     | IpfcModel.Display() |     | beforecalling |     |
| ----------------- | --- | ------------------- | --- | ------------- | --- |
IpfcModel.Export().
| Exporting | Using | Coordinate | System |     |     |
| --------- | ----- | ---------- | ------ | --- | --- |
Themethodsand propertiesdescribed in thissectionsupport theexportoffiles
withinformationabout thefacetedsolidmodels(withoutdatumsandsurfaces).
Thefilesareexportedin referenceto thecoordinate-systemfeatureinthe model
beingexported.
MethodsandPropertiesIntroduced:
• IpfcCoordSysExportInstructions.CsysName
• IpfcCoordSysExportInstructions.Quality
• IpfcCoordSysExportInstructions.MaxChordHeight
• IpfcCoordSysExportInstructions.AngleControl
• IpfcCoordSysExportInstructions.GetSliceExportData()
• IpfcCoordSysExportInstructions.SetSliceExportData()
• IpfcCoordSysExportInstructions.StepSize
• IpfcCoordSysExportInstructions.FacetControlOptions
• IpfcSelection.SetIntf3DCsys()
| ThepropertyIpfcCoordSysExportInstructions.CsysName |     |     |     |     | returns |
| -------------------------------------------------- | --- | --- | --- | --- | ------- |
thenameof thethename ofacoordinatesystemfeaturein the modelbeing
exported.Itis recommendedtousethe coordinatesystem thatplacesthe partor
assemblyin itsupper-rightquadrant,sothatall positionanddistance valuesof the
| exportedassembly | orpartare | positive. |                                 |     |     |
| ---------------- | --------- | --------- | ------------------------------- | --- | --- |
| 282              |           |           | Creo®ParametricVBAPIUser’sGuide |     |     |

| ThepropertyIpfcCoordSysExportInstructions.Quality |     |     | canbe |
| ------------------------------------------------- | --- | --- | ----- |
usedinsteadof
| IpfcCoordSysExportInstructions.MaxChordHeight |     |     | and |
| --------------------------------------------- | --- | --- | --- |
IpfcCoordSysExportInstructions.AngleControl.Youcan setthe
valuebetween1and10. Thehigher thevalue youpass,theloweris theMaximum
ChordHeight settingandhigher istheAngle Controlsettingthemethoduses. The
| defaultQualityvalueis | 1.0. |     |     |
| --------------------- | ---- | --- | --- |
Note
| Youmustpassthevalue | ofQualityas | NULL,if youuseMaximumChord |     |
| ------------------- | ----------- | -------------------------- | --- |
Height andAngleControlvalues.IfQuality,MaximumChordHeight, and
| AngleControlareall | NULL,thenthe | Qualitysettingof3is | used. |
| ------------------ | ------------ | ------------------- | ----- |
Usethe property
| IpfcCoordSysExportInstructions.MaxChordHeight |     |     | towork with |
| --------------------------------------------- | --- | --- | ----------- |
themaximumchordheightfor theexported files.The defaultvalueis 0.1.
Note
Youmustpassthevalue ofMaximumChordHeight asNULL, ifyouspecify
| theQuality value. |     |     |     |
| ----------------- | --- | --- | --- |
ThepropertyIpfcCoordSysExportInstructions.AngleControl
allowyoutowork withthe anglecontrolsettingfor theexportedfiles. Thedefault
valueis0.1.
Note
Youmustpassthevalue ofAngleControlvalueas NULL,if youspecify the
Qualityvalue.
Themethods
IpfcCoordSysExportInstructions.GetSliceExportData() and
IpfcCoordSysExportInstructions.SetSliceExportData()
get
andsettheIpfcSliceExportData dataobjectthatspecifiesdatafor theslice
| export.Theoptions | in thisobjectaredescribed | asfollows: |     |
| ----------------- | ------------------------- | ---------- | --- |
• CompIds—Specifiesthesequence ofintegersthatidentify thecomponentsthat
formthe pathfrom theroot assemblydown tothe componentpartorassembly
beingreferredto.Use thepropertyIpfcSliceExportData.CompIds to
| workwiththe | componentIDs. |     |     |
| ----------- | ------------- | --- | --- |
Interface 283

ThepropertyIpfcCoordSysExportInstructions.StepSize controls
| thestepsize | forthe | exportedfiles. | Thedefaultvalue |     | is0.0. |     |
| ----------- | ------ | -------------- | --------------- | --- | ------ | --- |
Note
Youmustpassthevalue ofStep Sizevalue asNULL, ifyouspecify the
Qualityvalue.
Theproperty
IpfcCoordSysExportInstructions.FacetControlOptionscontrol
thefacetexport optionsusing bitflags. You cansetthebit flagsusing the
IpfcFacetControlFlagobject.Formoreinformation onthe bitflag values,
| pleaserefer | tothe                                         | sectionExportingto |     | FacetedFormatson |     | page280. |
| ----------- | --------------------------------------------- | ------------------ | --- | ---------------- | --- | -------- |
| Thefunction | IpfcSelection.SetIntf3DCsys()setsthereference |                    |     |                  |     |          |
coordinatesystem forthe export.Theinput argumentReferenceCsys isthe
referencecoordinatesystemselection.Call thismethodwithoutanyargumentto
setdefault coordinatesystem. Referencecoordinatesystem isnotsupported for
| CADDSandNEUTRAL |     | filetypes. |     |     |     |     |
| --------------- | --- | ---------- | --- | --- | --- | --- |
| Exporting       |     | to PDF     | and | U3D |     |     |
Themethodsand propertiesdescribed in thissectionsupport theexportof Creo
| Parametricdrawingsand |     | solidmodelsto |     | PortableDocumentFormat |     |     |
| --------------------- | --- | ------------- | --- | ---------------------- | --- | --- |
(PDF)andU3Dformat. Youcanexport adrawing ora 2Dmodelasa 2Draster
imageembedded ina PDFfile. You canexportCreo Parametricsolidmodelsin
thefollowingways:
| • AsaU3D | modelembedded |     | ina | one-pagePDF | file |     |
| -------- | ------------- | --- | --- | ----------- | ---- | --- |
• As2Drasterimagesembedded inthe pagesofaPDF filerepresentingsaved
views
| • Asastandalone |     | U3Dfile |     |     |     |     |
| --------------- | --- | ------- | --- | --- | --- | --- |
Whileexportingmultiple sheetsofa CreoParametricdrawing to aPDF file,you
| canchooseto | exportall | sheets,thecurrentsheet,or |     |     | selectedsheets. |     |
| ----------- | --------- | ------------------------- | --- | --- | --------------- | --- |
Thesemethodsalsoallow youtoinsert avarietyof non-geometricinformationto
| improvedocumentcontent,navigation, |     |     |     | andsearch. |     |     |
| ---------------------------------- | --- | --- | --- | ---------- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcPDFExportInstructions.Create()
• IpfcPDFExportInstructions.FilePath
• IpfcPDFExportInstructions.Options
• IpfcPDFExportInstructions.ProfilePath
| 284 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

• CCpfcPDFOption.Create()
• IpfcPDFOption.OptionType
• PDFOption.OptionValue
| ThemethodCCpfcPDFExportInstructions.Create() |     |     |     |     |     |     |     |     | createsanew |     |
| -------------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- | ----------- | --- |
theIpfcPDFExportInstructionsdataobjectthatdescribes
instanceof
howto export CreoParametricdrawings orsolidmodelsto thePDF andU3D
| formats.The |                                    | optionsin | this | objectare | describedas |             | follows: |                 |     |     |
| ----------- | ---------------------------------- | --------- | ---- | --------- | ----------- | ----------- | -------- | --------------- | --- | --- |
| •           | FilePath—Specifiesthe              |           |      | nameofthe |             | outputfile. | Usethe   | property        |     |     |
|             | IpfcPDFExportInstructions.FilePath |           |      |           |             |             |          | to setthenameof |     | the |
outputfile.
| •   | Options—Specifiesa |     | collectionofPDF |     |     | exportoptions |     | ofthe | type |     |
| --- | ------------------ | --- | --------------- | --- | --- | ------------- | --- | ----- | ---- | --- |
IpfcPDFOption.Createa newinstanceof thisobjectusing themethod
IpfcPDFOption.Create().This objectcontainsthe followingattributes:
|     | ○ OptionType—Specifiesthetypeofoption |     |     |                  |     |     | in termsofthe |     |          |     |
| --- | ------------------------------------- | --- | --- | ---------------- | --- | --- | ------------- | --- | -------- | --- |
|     | EpfcPDFOptionType                     |     |     | enumeratedclass. |     |     | Setthisoption |     | usingthe |     |
propertyIpfcPDFOption.OptionType.
|     | ○ OptionValue—Specifiesthe       |     |     |     | valueofthe |     | optionin          | termsof | the |     |
| --- | -------------------------------- | --- | --- | --- | ---------- | --- | ----------------- | ------- | --- | --- |
|     | IpfcArgValueobject.Setthisoption |     |     |     |            |     | using theproperty |         |     |     |
IpfcPDFOption.OptionValue.
|     | Usethe                                | propertyIpfcPDFExportInstructions.Options |                |     |              |     |     |             |              | tosetthe |
| --- | ------------------------------------- | ----------------------------------------- | -------------- | --- | ------------ | --- | --- | ----------- | ------------ | -------- |
|     | collectionofPDF                       |                                           | exportoptions. |     |              |     |     |             |              |          |
| •   | ProfilePath—Specifiestheexport        |                                           |                |     | profilepath. |     | Use | theproperty |              |          |
|     | IpfcPDFExportInstructions.ProfilePath |                                           |                |     |              |     |     | tosetthe    | profilepath. |          |
Whenyousettheprofilepath, thePDF exportoptions setinthe dataobject
|     | IpfcPDFExportInstructionsdataobjectareignoredwhenthe |     |     |            |     |                             |     |     |     | method |
| --- | ---------------------------------------------------- | --- | --- | ---------- | --- | --------------------------- | --- | --- | --- | ------ |
|     | IpfcModel.Export()is                                 |     |     | called.You |     | cansettheprofilepathasNULL. |     |     |     |        |
Note
|     | Youcan | specifythe | profilepathonly |     |     | for | drawings. |     |     |     |
| --- | ------ | ---------- | --------------- | --- | --- | --- | --------- | --- | --- | --- |
Thetypesofoptions (givenby theEpfcPDFOptionType enumerated class)
| availablefor |                               | exportto | PDFand | U3Dformatsaredescribed |     |     |             | asfollows:   |     |     |
| ------------ | ----------------------------- | -------- | ------ | ---------------------- | --- | --- | ----------- | ------------ | --- | --- |
| •            | EpfcPDFOPT_FONT_STROKE—Allows |          |        |                        |     |     | youtoswitch | betweenusing |     |     |
TrueTypefontsor “stroking”textinthe resulting document.Thisoption is
|     | givenby | theEpfcPDFFontStrokeMode |     |     |     |     | enumeratedclass |     | andtakesthe |     |
| --- | ------- | ------------------------ | --- | --- | --- | --- | --------------- | --- | ----------- | --- |
followingvalues:
| Interface |     |     |     |     |     |     |     |     |     | 285 |
| --------- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |

○ EpfcPDF_USE_TRUE_TYPE_FONTS—SpecifiesTrueTypefonts.
|     | This                                 | isthedefaulttype. |     |     |     |     |           |             |     |
| --- | ------------------------------------ | ----------------- | --- | --- | --- | --- | --------- | ----------- | --- |
|     | ○ EpfcPDF_STROKE_ALL_FONTS—Specifies |                   |     |     |     |     | theoption | tostrokeall |     |
fonts.
| •   | EpfcPDFOPT_COLOR_DEPTH—Allows |                       |     |     | youtochoose |           | betweencolor, |     |     |
| --- | ----------------------------- | --------------------- | --- | --- | ----------- | --------- | ------------- | --- | --- |
|     | grayscale,or                  | monochromeoutput.This |     |     | option      | isgivenby |               | the |     |
EpfcPDFColorDepth
|     |     |     | enumeratedclass |     |     | andtakesthefollowingvalues: |     |     |     |
| --- | --- | --- | --------------- | --- | --- | --------------------------- | --- | --- | --- |
○ EpfcPDF_CD_COLOR—Specifiescoloroutput. Thisis thedefaultvalue.
|     | ○ EpfcPDF_CD_GRAY—Specifies             |     |     |     | grayscaleoutput.  |     |                |     |     |
| --- | --------------------------------------- | --- | --- | --- | ----------------- | --- | -------------- | --- | --- |
|     | ○ EpfcPDF_CD_MONO—Specifies             |     |     |     | monochromeoutput. |     |                |     |     |
| •   | EpfcPDFOPT_HIDDENLINE_MODE—Enablesyouto |     |     |     |                   |     | setthestylefor |     |     |
hiddenlinesin theresulting PDF document.Thisoption isgivenby the
|     | EpfcPDFHiddenLineMode |     |     | enumeratedclass |     | andtakesthefollowing |     |     |     |
| --- | --------------------- | --- | --- | --------------- | --- | -------------------- | --- | --- | --- |
values:
|     | ○ EpfcPDF_HLM_SOLID—Specifies        |     |     |     | solidhiddenlines. |              |     |      |       |
| --- | ------------------------------------ | --- | --- | --- | ----------------- | ------------ | --- | ---- | ----- |
|     | ○ EpfcPDF_HLM_DASHED—Specifiesdashed |     |     |     |                   | hiddenlines. |     | This | isthe |
defaulttype.
• EpfcPDFOPT_SEARCHABLE_TEXT—If true,stroked textissearchable.
|     | Thedefaultvalue                   | istrue. |     |     |     |                     |     |     |     |
| --- | --------------------------------- | ------- | --- | --- | --- | ------------------- | --- | --- | --- |
| •   | EpfcPDFOPT_RASTER_DPI—Allowsyouto |         |     |     |     | settheresolutionfor |     |     | the |
outputof anyshadedviews in DPI.Itcan takeavaluebetween 100and600.
|     | Thedefaultvalue             | is300.         |       |     |               |     |                 |     |     |
| --- | --------------------------- | -------------- | ----- | --- | ------------- | --- | --------------- | --- | --- |
| •   | EpfcPDFOPT_LAUNCH_VIEWER—If |                |       |     | true,launches |     | theAdobeAcrobat |     |     |
|     | Reader.The                  | defaultvalueis | true. |     |               |     |                 |     |     |
• EpfcPDFOPT_LAYER_MODE—Enablesyouto settheavailability oflayers
inthe document.Itis givenby theEpfcPDFLayerModeenumerated class
andtakesthefollowingvalues:
○ EpfcPDF_LAYERS_ALL—Exportsthe visiblelayersandentities. Thisis
thedefault.
|     | ○ EpfcPDF_LAYERS_VISIBLE—Exports |     |     |     |     | onlyvisiblelayers |     |     | ina |
| --- | -------------------------------- | --- | --- | --- | --- | ----------------- | --- | --- | --- |
drawing.
|     | ○ EpfcPDF_LAYERS_NONE—Exports |                |     |                     | only | thevisibleentitiesin |     |     | the |
| --- | ----------------------------- | -------------- | --- | ------------------- | ---- | -------------------- | --- | --- | --- |
|     | drawing,but                   | notthelayerson |     | whichtheyareplaced. |      |                      |     |     |     |
• EpfcPDFOPT_PARAM_MODE—Enablesyouto settheavailability ofmodel
parametersas searchablemetadatain thePDFdocument.It isgiven bythe
|     | EpfcPDFParameterModeenumerated |     |     |     | classandtakesthe |     |     | following |     |
| --- | ------------------------------ | --- | --- | --- | ---------------- | --- | --- | --------- | --- |
values:
| 286 |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- | --- | --- |

| ○ EpfcPDF_PARAMS_ALL—Exportsthe     |              |                |     | drawing    | andthemodel            |               |     |
| ----------------------------------- | ------------ | -------------- | --- | ---------- | ---------------------- | ------------- | --- |
| parameterstoPDF.This                |              | isthe default. |     |            |                        |               |     |
| ○ EpfcPDF_PARAMS_DESIGNATED—Exports |              |                |     |            | only thespecifiedmodel |               |     |
| parametersinthe                     | PDFmetadata. |                |     |            |                        |               |     |
| ○ EpfcPDF_PARAMS_NONE—Exports       |              |                |     | thedrawing | to                     | PDFwithoutthe |     |
modelparameters.
• EpfcPDFOPT_HYPERLINKS—Setshyperlinksto be exportedas labeltext
onlyor setstheunderlyinghyperlinkURLs asactive.The defaultvalueis true,
specifyingthatthehyperlinksareactive.
| • EpfcPDFOPT_BOOKMARK_ZONES—Iftrue, |     |     |     | addsbookmarksto |     |     | thePDF |
| ----------------------------------- | --- | --- | --- | --------------- | --- | --- | ------ |
showingzoomedin regions orzonesin thedrawing sheet.The zoneonan A4-
sizedrawing sheetisignored.
| • EpfcPDFOPT_BOOKMARK_VIEWS—Iftrue,   |              |            |                | addsbookmarksto |             |                | thePDF |
| ------------------------------------- | ------------ | ---------- | -------------- | --------------- | ----------- | -------------- | ------ |
| documentshowingzoomed                 |              | inviews    | on thedrawing. |                 |             |                |        |
| • EpfcPDFOPT_BOOKMARK_SHEETS—If       |              |            |                | true,adds       | bookmarksto |                | thePDF |
| documentshowingeachof                 |              | thedrawing | sheets.        |                 |             |                |        |
| • EpfcPDFOPT_BOOKMARK_FLAG_NOTES—If   |              |            |                | true,adds       |             | bookmarkstothe |        |
| PDFdocumentshowing                    | thetextofthe |            | flagnote.      |                 |             |                |        |
| • EpfcPDFOPT_TITLE—Specifiesatitlefor |              |            |                | thePDFdocument. |             |                |        |
• EpfcPDFOPT_AUTHOR—Specifies thenameof thepersongeneratingthe
PDFdocument.
| • EpfcPDFOPT_SUBJECT—Specifiesthe |     |     |                    | subjectofthePDF |     | document. |     |
| --------------------------------- | --- | --- | ------------------ | --------------- | --- | --------- | --- |
| • EpfcPDFOPT_KEYWORDS—Specifies   |     |     | relevantkeywordsin |                 |     | thePDF    |     |
document.
| • EpfcPDFOPT_PASSWORD_TO_OPEN—Sets |     |     |     | apasswordto |     | open | thePDF |
| ---------------------------------- | --- | --- | --- | ----------- | --- | ---- | ------ |
document.Bydefault,this option isNULL,which meansanyone canopen the
| PDFdocumentwithouta | password. |     |     |     |     |     |     |
| ------------------- | --------- | --- | --- | --- | --- | --- | --- |
EpfcPDFOPT_MASTER_PASSWORD—Sets
| •   |     |     |     | apassword |     | torestrict | orlimit |
| --- | --- | --- | --- | --------- | --- | ---------- | ------- |
theoperationsthattheviewer canperformon theopened PDFdocument.By
default,thisoption isNULL, whichmeansyoucan makeanychangesto the
| PDFdocumentregardless | ofthe | settingsofthemodification |     |     |     | flags |     |
| --------------------- | ----- | ------------------------- | --- | --- | --- | ----- | --- |
EpfcPDFOPT_ALLOW_*.
• EpfcPDFOPT_RESTRICT_OPERATIONS—If true,enables youtorestrict
thePDF
| orlimitoperationson |     | document.Bydefault,is |     |     | isfalse. |     |     |
| ------------------- | --- | --------------------- | --- | --- | -------- | --- | --- |
• EpfcPDFOPT_ALLOW_MODE—Enablesyouto setthesecurity settingsfor
| thePDFdocument.This | option  | mustbesetif |     | PDFOPT_RESTRICT_ |     |     |     |
| ------------------- | ------- | ----------- | --- | ---------------- | --- | --- | --- |
| OPERATIONSissetto   | true.It | isgivenby   |     | the              |     |     |     |
| Interface           |         |             |     |                  |     |     | 287 |

|     | EpfcPDFRestrictOperationsModeenumerated |     |     |     |     | classandtakesthe |     |     |
| --- | --------------------------------------- | --- | --- | --- | --- | ---------------- | --- | --- |
followingvalues:
|     | ○ EpfcPDF_RESTRICT_NONE—Specifies |     |     |     | thattheusercan |     | performany |     |
| --- | --------------------------------- | --- | --- | --- | -------------- | --- | ---------- | --- |
thePDFdocument.This
|     | ofthe permittedviewer |     | operationson |     |     |     | isthe |     |
| --- | --------------------- | --- | ------------ | --- | --- | --- | ----- | --- |
defaultvalue.
|     | ○ EpfcPDF_RESTRICT_FORMS_SIGNING—Restricts        |                     |                 |            |              |                 | theuserfrom |     |
| --- | ------------------------------------------------- | ------------------- | --------------- | ---------- | ------------ | --------------- | ----------- | --- |
|     | adding                                            | digitalsignaturesto | thePDFdocument. |            |              |                 |             |     |
|     | ○ EpfcPDF_RESTRICT_INSERT_DELETE_ROTATE—Restricts |                     |                 |            |              |                 | theuser     |     |
|     | from inserting,deleting,orrotating                |                     |                 | thepagesin |              | thePDFdocument. |             |     |
|     | ○ EpfcPDF_RESTRICT_COMMENT_FORM_SIGNING—Restricts |                     |                 |            |              |                 | theuser     |     |
|     | from addingor                                     | editingcomments     |                 | inthe      | PDFdocument. |                 |             |     |
|     | ○ EpfcPDF_RESTRICT_EXTRACTING—Restrictsthe        |                     |                 |            |              |                 | userfrom    |     |
PDFdocument.
extractingpagesfromthe
| •   | EpfcPDFOPT_ALLOW_PRINTING—Iftrue,           |     |       |     | allowsyouto |     | printthePDF |     |
| --- | ------------------------------------------- | --- | ----- | --- | ----------- | --- | ----------- | --- |
|     | document.Bydefault,itis                     |     | true. |     |             |     |             |     |
| •   | EpfcPDFOPT_ALLOW_PRINTING_MODE—Enablesyouto |     |       |     |             |     | settheprint |     |
resolution.Itisgiven bythe EpfcPDFPrintingMode enumeratedclass
andtakesthefollowingvalues:
EpfcPDF_PRINTING_LOW_RES—Specifies
|     | ○   |     |     |     |     | lowresolutionfor |     |     |
| --- | --- | --- | --- | --- | --- | ---------------- | --- | --- |
printing.
|     | ○ EpfcPDF_PRINTING_HIGH_RES—Specifies |              |        |     |     | high | resolutionfor |     |
| --- | ------------------------------------- | ------------ | ------ | --- | --- | ---- | ------------- | --- |
|     | printing.This                         | isthedefault | value. |     |     |      |               |     |
• EpfcPDFOPT_ALLOW_COPYING—If true,allowsyou tocopy contentfrom
thePDFdocument.Bydefault,itistrue.
| •   | EpfcPDFOPT_ALLOW_ACCESSIBILITY—If |     |     |     |     | true,enables | visually- |     |
| --- | --------------------------------- | --- | --- | --- | --- | ------------ | --------- | --- |
impairedscreenreader devicestoextract dataindependentofthe valuegiven
|     | bythe EpfcPDFRestrictOperationsModeenumeratedclass. |       |     |                     |     |     | The        |     |
| --- | --------------------------------------------------- | ----- | --- | ------------------- | --- | --- | ---------- | --- |
|     | defaultvalueis                                      | true. |     |                     |     |     |            |     |
| •   | EpfcPDFOPT_PENTABLE—Iftrue,                         |       |     | usesthestandardCreo |     |     | Parametric |     |
pentabletocontrol thelineweight,linestyle, andlinecolorofthe exported
|     | geometry.The | defaultvalueis | false. |     |     |     |     |     |
| --- | ------------ | -------------- | ------ | --- | --- | --- | --- | --- |
• EpfcPDFOPT_LINECAP—Enables youto controlthetreatmentofthe ends
ofthegeometry linesexportedto PDF.Itis givenby theEpfcPDFLinecap
|     | enumeratedclass                  | andtakesthefollowingvalues: |     |     |                          |     |     |     |
| --- | -------------------------------- | --------------------------- | --- | --- | ------------------------ | --- | --- | --- |
|     | ○ EpfcPDF_LINECAP_BUTT—Specifies |                             |     |     | thebuttcapsquareend.This |     |     | is  |
thedefaultvalue.
|     | ○ EpfcPDF_LINECAP_ROUND—Specifies |     |     |     | theround |                                 | capend. |     |
| --- | --------------------------------- | --- | --- | --- | -------- | ------------------------------- | ------- | --- |
| 288 |                                   |     |     |     |          | Creo®ParametricVBAPIUser’sGuide |         |     |

○ EpfcPDF_LINECAP_PROJECTING_SQUARE—Specifies the
projectingsquarecapend.
• EpfcPDFOPT_LINEJOIN—Enablesyouto controlthetreatmentof the
joinedcornersofconnected linesexportedto PDF.Itis givenby the
EpfcPDFLinejoin enumeratedclass andtakesthefollowingvalues:
○ EpfcPDF_LINEJOIN_MITER—Specifies themiterjoin.This isthe
default.
○ EpfcPDF_LINEJOIN_ROUND—Specifies theround join.
○ EpfcPDF_LINEJOIN_BEVEL—Specifies thebeveljoin.
• EpfcPDFOPT_SHEETS—Allows youto specifythe sheetsfroma Creo
Parametricdrawingthatare tobe exported toPDF.It isgivenby the
EpfcPrintSheets enumeratedclass andtakesthefollowingvalues:
○ EpfcPRINT_CURRENT_SHEET—Only thecurrentsheetisexported to
PDF
○ EpfcPRINT_ALL_SHEETS—All thesheetsare exportedto PDF.Thisis
thedefaultvalue.
○ EpfcPRINT_SELECTED_SHEETS—Sheets ofa specifiedrangeare
exportedto PDF.Ifthis valueis assigned,thenthevalue oftheoption
EpfcPDFOPT_SHEET_RANGE mustalsobe known.
• EpfcPDFOPT_SHEET_RANGE—Specifiesthe rangeofsheetsin adrawing
thatareto beexported to PDF.If thisoption isset, thenthe option
EpfcPDFOPT_SHEETS mustbesetto thevalue EpfcPRINT_SELECTED_
SHEETS.
• EpfcPDFOPT_EXPORT_MODE—Enables youto selectthe objectto be
exportedto PDFandthe exportformat. Itis givenby the
EpfcPDFExportMode enumeratedclass andtakesthefollowingvalues:
○ EpfcPDF_2D_DRAWING—Onlydrawings areexported toPDF.This is
thedefaultvalue.
○ EpfcPDF_3D_AS_NAMED_VIEWS—3D modelsareexportedas 2D
rasterimagesembeddedin PDFfiles.
○ EpfcPDF_3D_AS_U3D_PDF—3D modelsareexportedas U3Dmodels
embeddedin one-pagePDF files.
○ EpfcPDF_3D_AS_U3D—A 3D modelis exportedas aU3D(.u3d) file.
This valueignoresthe options setfor theEpfcPDFOptionType
enumeratedclass.
• EpfcPDFOPT_LIGHT_DEFAULT—Enablesyouto setthedefaultlighting
styleusedwhileexporting3D modelsin the U3Dformattoa one-pagePDF
file,thatiswhen theoption EpfcPDFOPT_EXPORT_MODE issetto
Interface 289

| EpfcPDF_3D_AS_U3D.The            |     |     | valuesfor | thisoption |        | aregivenby |     | the |
| -------------------------------- | --- | --- | --------- | ---------- | ------ | ---------- | --- | --- |
| EpfcPDFU3DLightingModeenumerated |     |     |           |            | class. |            |     |     |
• EpfcPDFOPT_RENDER_STYLE_DEFAULT—Enables youto setthedefault
renderingstyleused whileexportingCreoParametricmodelsin theU3D
| formattoa | one-pagePDFfile, |     | thatiswhen |     | theoption | EpfcPDFOPT_ |     |     |
| --------- | ---------------- | --- | ---------- | --- | --------- | ----------- | --- | --- |
EXPORT_MODEis setto EpfcPDF_3D_AS_U3D.The valuesfor thisoption
| aregivenby | theEpfcPDFU3DRenderMode |     |     |     | enumeratedclass. |     |     |     |
| ---------- | ----------------------- | --- | --- | --- | ---------------- | --- | --- | --- |
• EpfcPDFOPT_SIZE—Allows youto specifythe pagesize ofthe exported
PDFfile. Thevalues forthis optionaregiven bytheEpfcPlotPaperSize
enumeratedclass.Ifthe valueis settoEpfcVARIABLESIZEPLOT, youalso
| needtosetthe | optionsEpfcPDFOPT_HEIGHT |     |     |     |     | andEpfcPDFOPT_WIDTH. |     |     |
| ------------ | ------------------------ | --- | --- | --- | --- | -------------------- | --- | --- |
• EpfcPDFOPT_HEIGHT—Enablesyouto settheheightfor auser-defined
exportedPDF
| pagesizeofthe |     |     | file.The | defaultvalueis |     |     | 0.0. |     |
| ------------- | --- | --- | -------- | -------------- | --- | --- | ---- | --- |
• EpfcPDFOPT_WIDTH—Enables youtosetthe widthfora user-definedpage
| sizeof | theexportedPDF |     | file.The defaultvalueis |     |     | 0.0. |     |     |
| ------ | -------------- | --- | ----------------------- | --- | --- | ---- | --- | --- |
• EpfcPDFOPT_ORIENTATION—Enables youto specifythe orientationof
| thepagesin                     | theexportedPDF |     | file.It isgivenby |        |     | the |     |     |
| ------------------------------ | -------------- | --- | ----------------- | ------ | --- | --- | --- | --- |
| EpfcSheetOrientationenumerated |                |     |                   | class. |     |     |     |     |
EpfcORIENT_PORTRAIT—Exports
| ○                              |                    |     |     |     | thepagesin          |     | portraitorientation. |     |
| ------------------------------ | ------------------ | --- | --- | --- | ------------------- | --- | -------------------- | --- |
| This                           | isthedefaultvalue. |     |     |     |                     |     |                      |     |
| ○ EpfcORIENT_LANDSCAPE—Exports |                    |     |     |     | thepagesinlandscape |     |                      |     |
orientation.
• EpfcPDFOPT_TOP_MARGIN—Allowsyouto specify thetopmargin ofthe
| viewport.The | defaultvalue |     | is0.0. |     |     |     |     |     |
| ------------ | ------------ | --- | ------ | --- | --- | --- | --- | --- |
• EpfcPDFOPT_LEFT_MARGIN—Allows youtospecify theleftmargin of
| theview                                        | port.Thedefault |     | valueis0.0. |     |     |     |     |            |
| ---------------------------------------------- | --------------- | --- | ----------- | --- | --- | --- | --- | ---------- |
| • EpfcPDFOPT_BACKGROUND_COLOR_RED—Specifiesthe |                 |     |             |     |     |     |     | defaultred |
U3Dmodel.You
| backgroundcolorthatappears |              |     | behindthe              |     |     |      | cansetanyvalue |     |
| -------------------------- | ------------ | --- | ---------------------- | --- | --- | ---- | -------------- | --- |
| withinthe                  | rangeof0.0to |     | 1.0.The defaultvalueis |     |     | 1.0. |                |     |
• EpfcPDFOPT_BACKGROUND_COLOR_GREEN—Specifiesthe defaultgreen
backgroundcolorthatappears behindthe U3Dmodel.You cansetanyvalue
| withinthe                                    | rangeof0.0to |     | 1.0.The defaultvalueis |     |     | 1.0. |                |     |
| -------------------------------------------- | ------------ | --- | ---------------------- | --- | --- | ---- | -------------- | --- |
| • EpfcPDFOPT_BACKGROUND_COLOR_BLUE—Specifies |              |     |                        |     |     |      | thedefaultblue |     |
backgroundcolorthatappears behindthe U3Dmodel.You cansetanyvalue
| withinthe | rangeof0.0to |     | 1.0.The defaultvalueis |     |     | 1.0. |     |     |
| --------- | ------------ | --- | ---------------------- | --- | --- | ---- | --- | --- |
• EpfcPDFOPT_ADD_VIEWS—If true,allowsyouto addview definitionsto
| theU3Dmodelfrom |     | afile. | Bydefault,itis |     | true. |                                 |     |     |
| --------------- | --- | ------ | -------------- | --- | ----- | ------------------------------- | --- | --- |
| 290             |     |        |                |     |       | Creo®ParametricVBAPIUser’sGuide |     |     |

| •   | EpfcPDFOPT_VIEW_TO_EXPORT—Specifies |     |     |     |     | theview | orviews | to be |
| --- | ----------------------------------- | --- | --- | --- | --- | ------- | ------- | ----- |
exportedto thePDF file.It isgivenby theEpfcPDFSelectedViewMode
|     | enumeratedclass | andtakesthefollowingvalues: |     |     |     |     |     |     |
| --- | --------------- | --------------------------- | --- | --- | --- | --- | --- | --- |
EpfcPDF_VIEW_SELECT_CURRENT—Exportsthe
|     | ○       |             |       |     |     |     | currentgraphical |     |
| --- | ------- | ----------- | ----- | --- | --- | --- | ---------------- | --- |
|     | areatoa | one-pagePDF | file. |     |     |     |                  |     |
○ EpfcPDF_VIEW_SELECT_ALL—Exports all theviews toa multi-page
PDFfile. Eachpagecontainsoneview with theview namedisplayedatthe
|     | bottomcenterofthe                        |     | viewport. |     |     |     |              |      |
| --- | ---------------------------------------- | --- | --------- | --- | --- | --- | ------------ | ---- |
|     | ○ EpfcPDF_VIEW_SELECT_BY_NAME—Exportsthe |     |           |     |     |     | selectedview | to a |
one-pagePDF filewiththeview nameprinted atthe bottomcenterofthe
|     | view port.Ifthisvalue         |          | isassigned,                  |      | thenthe   | optionEpfcPDFOPT_ |     |       |
| --- | ----------------------------- | -------- | ---------------------------- | ---- | --------- | ----------------- | --- | ----- |
|     | SELECTED_VIEW                 |          | mustalsobe                   | set. |           |                   |     |       |
| •   | EpfcPDFOPT_SELECTED_VIEW—Sets |          |                              |      | theoption | EpfcPDFOPT_VIEW_  |     |       |
|     | TO_EXPORTto                   | thevalue | EpfcPDF_VIEW_SELECT_BY_NAME, |      |           |                   |     | ifthe |
|     | correspondingviewis           |          | successfullyfound.           |      |           |                   |     |       |
• EpfcPDFOPT_PDF_SAVE—Specifies thePDF saveoptions.It isgiven by
theEpfcPDFSaveMode enumeratedclass andtakesthefollowingvalues:
|     | ○ EpfcPDF_ARCHIVE_1—Applicableonly |     |                         |       |      | forthe                  | valueEpfcPDF_2D_ |     |
| --- | ---------------------------------- | --- | ----------------------- | ----- | ---- | ----------------------- | ---------------- | --- |
|     | DRAWING.Saves                      |     | thedrawings             | asPDF | with | thefollowingconditions: |                  |     |
|     | ◆ The valueof                      |     | EpfcPDFLayerModeissetto |       |      | EpfcPDF_LAYERS_         |                  |     |
NONE.
|     |               |     | EpfcPDFOPT_HYPERLINKS |     |     |         | FALSE. |     |
| --- | ------------- | --- | --------------------- | --- | --- | ------- | ------ | --- |
|     | ◆ The valueof |     |                       |     |     | issetto |        |     |
◆ The shadedviews in thedrawings willnot havetransparencyand may
|     | overlapotherdatain |     | thePDF.                            |     |     |     |       |       |
| --- | ------------------ | --- | ---------------------------------- | --- | --- | --- | ----- | ----- |
|     | ◆ The valueof      |     | EpfcPDFOPT_PASSWORD_TO_OPENissetto |     |     |     |       | NULL. |
|     | ◆ The valueof      |     | EpfcPDFOPT_MASTER_PASSWORDis       |     |     |     | setto | NULL. |
○ EpfcPDF_FULL—SavesthePDF with thevaluessetby you. Thisis the
defaultvalue.
| Exporting |     | 3D  | Geometry |     |     |     |     |     |
| --------- | --- | --- | -------- | --- | --- | --- | --- | --- |
TheVB APIallowsyou toexport threedimensionalgeometryto various formats.
| Export | Instructions |     |     |     |     |     |     |     |
| ------ | ------------ | --- | --- | --- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcModel.ExportIntf3D()
• IpfcBaseSession.ExportProfileLoad()
| Interface |     |     |     |     |     |     |     | 291 |
| --------- | --- | --- | --- | --- | --- | --- | --- | --- |

• CCpfcGeometryFlags.Create()
• CCpfcInclusionFlags.Create()
• CCpfcLayerExportOptions.Create()
• CCpfcTriangulationInstructions.Create()
FromCreo Parametric5.0F000onward,the followinginterfaces alongwiththeir
methodshavebeendeprecated. Use themethod
IpfcModel.ExportIntf3D()insteadto export CreoParametricmodelsto
otherfileformats. Alltheoptions thatcanbe setwith theseinterfacesand
methods,canalsobe setusing theexportprofileoption in CreoParametric.Refer
tothe CreoParametricDataExchangeOnline Helpfor moreinformation.
• Export3DInstructions
• ACIS3DExportInstructions
• CATIAModel3DExportInstructions
• CATIASession3DExportInstructions
• CatiaPart3DExportInstructions
• CatiaProduct3DExportInstructions
• CatiaCGR3DExportInstructions
• DXF3DExportInstructions
• DWG3DExportInstructions
• IGES3DNewExportInstructions
• JT3DExportInstructions
• ParaSolid3DExportInstructions
• STEP3DExportInstructions
• SWPart3DExportInstructions
• SWAsm3DExportInstructions
• UG3DExportInstructions
• VDA3DExportInstructions
ThemethodIfcModel.ExportIntf3D() exportsaCreo Parametricmodel
tothe specifiedoutput formatusing thedefaultexport profile.The exportoptions
mustbe setusing theexport profileoption in CreoParametric.
ThemethodIpfcBaseSession.ExportProfileLoad() loadsthe
specifiedprofilefor export.You canusethis functionwhen youwantto usethe
exportprofileofyourchoiceinsteadof thedefaultexport profileina particular
CreoParametricsession. Theinput argumentProfileFileisthe fullpathto the
profilealong withthe profilename andextension.
292 Creo®ParametricVBAPIUser’sGuide

Note
Once theexportprofilefileisloadedin aCreo Parametricsession, itwillbe
| activein | theinteractivemodeaswell. |     |     |     |
| -------- | ------------------------- | --- | --- | --- |
ThemethodCCpfcTriangulationInstructions.Create() createsa
| objectthatwillbe | usedto    | define theparametersforfacetedexports. |     |     |
| ---------------- | --------- | -------------------------------------- | --- | --- |
| Export           | Utilities |                                        |     |     |
MethodsIntroduced:
• IpfcBaseSession.IsConfigurationSupported()
• IpfcBaseSession.IsGeometryRepSupported()
ThemethodIpfcBaseSession.IsConfigurationSupported()
checkswhetherthespecifiedassembly configurationis validfor aparticular
modelandthe specifiedexportformat. Theinputparametersfor thismethodare:
• Configuration—Specifiesthestructureand contentofthe outputfiles.
| • Type—Specifiestheoutput |     | filetypeto | create. |     |
| ------------------------- | --- | ---------- | ------- | --- |
Themethodreturns atruevalue ifthe configurationis supportedfor thespecified
exporttype.
ThemethodIpfcBaseSession.IsGeometryRepSupported() checks
whetherthe specifiedgeometricrepresentation isvalidfor aparticularexport
| format.The        | inputparametersare | :                   |           |            |
| ----------------- | ------------------ | ------------------- | --------- | ---------- |
| • Flags—Thetypeof |                    | geometrysupportedby | theexport | operation. |
• Type—Theoutputfiletypeto create.
Themethodreturns atruevalue ifthe geometrycombinationisvalidfor the
| specifiedmodelandexport |     | type. |     |     |
| ----------------------- | --- | ----- | --- | --- |
ThemethodsIpfcBaseSession.IsConfigurationSupported() and
| IpfcBaseSession.IsGeometryRepSupported() |     |     |     | mustbe calledbefore |
| ---------------------------------------- | --- | --- | --- | ------------------- |
exportingan assembly tothe specifiedexportformatsexceptfor theCADDSand
STEP2Dformats. The returnvaluesof boththe methodsmustbe truefor the
| exportoperationto | be  | successful. |     |     |
| ----------------- | --- | ----------- | --- | --- |
Usethe methodIpfcModel.Export()to exportthe assemblyto thespecified
outputformat.
Interface 293

| Shrinkwrap | Export |     |     |     |
| ---------- | ------ | --- | --- | --- |
Toimprove performancein alarge assemblydesign,youcan exportlightweight
representationsofmodelscalledshrinkwrap models.Ashrinkwrap modelis based
ontheexternalsurfaces ofthesource partorassembly modelandcapturesthe
| outershapeof | thesourcemodel. |     |     |     |
| ------------ | --------------- | --- | --- | --- |
You cancreatethefollowingtypesofnonassociativeexported shrinkwrapmodels:
• SurfaceSubset—Thistypeconsists ofasubsetof theoriginal model’s
surfaces.
• FacetedSolid—Thistypeisa facetedsolidrepresentingtheoriginal solid.
• MergedSolid—Theexternalcomponentsfrom thereferenceassembly model
aremerged intoa singlepart representingthesolidgeometryin allcollected
components.
MethodsIntroduced:
• IpfcSolid.ExportShrinkwrap()
You canexportthe specifiedsolidmodelas ashrinkwrapmodelusing themethod
IpfcSolid.ExportShrinkwrap().This
methodtakesthe
| IpfcShrinkwrapExportInstructions |     |     | objectasan | argument. |
| -------------------------------- | --- | --- | ---------- | --------- |
Usethe appropriateinterface givenin thefollowingtableto createtherequired
typeof shrinkwrap.Alltheinterfaces havetheir ownstaticmethodto createan
objectof thespecifiedtype. Theobjectcreatedby theseinterfacescan beused as
| anobjectof | typeIpfcShrinkwrapExportInstructions |     |     | or  |
| ---------- | ------------------------------------ | --- | --- | --- |
IpfcShrinkwrapModelExportInstructions.
| TypeofShrinkwrapModel |     | InterfacetoUse                     |     |     |
| --------------------- | --- | ---------------------------------- | --- | --- |
| SurfaceSubset         |     | IpfcShrinkwrapSurfaceSubsetInstruc |     |     |
tions
| FacetedPart |     | IpfcShrinkwrapFacetedPartInstruc |     |     |
| ----------- | --- | -------------------------------- | --- | --- |
tions
| FacetedVRML |     | IpfcShrinkwrapFacetedVRMLInstruc |     |     |
| ----------- | --- | -------------------------------- | --- | --- |
tions
| FacetedSTL |     | IpfcShrinkwrapFacetedSTLInstruc |     |     |
| ---------- | --- | ------------------------------- | --- | --- |
tions
| MergedSolid |     | IpfcShrinkwrapMergedSolidInstruc |     |     |
| ----------- | --- | -------------------------------- | --- | --- |
tions
| Setting | Shrinkwrap | Options |     |     |
| ------- | ---------- | ------- | --- | --- |
TheinterfaceIpfcShrinkwrapModelExportInstructions contains the
generalmethodsavailableforall thetypesofshrinkwrap models.The object
createdbyany oftheinterfaces specifiedin the precedingtablecanbe usedwith
thesemethods.
PropertiesIntroduced:
| 294 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

• IpfcShrinkwrapModelExportInstructions.Method
• IpfcShrinkwrapModelExportInstructions.Quality
• IpfcShrinkwrapModelExportInstructions.AutoHoleFilling
• IpfcShrinkwrapModelExportInstructions.IgnoreSkeleton
• IpfcShrinkwrapModelExportInstructions.IgnoreQuilts
• IpfcShrinkwrapModelExportInstructions.AssignMassProperties
• IpfcShrinkwrapModelExportInstructions.IgnoreSmallSurfaces
• IpfcShrinkwrapModelExportInstructions.SmallSurfPercentage
• IpfcShrinkwrapModelExportInstructions.DatumReferences
ThepropertyIpfcShrinkwrapModelExportInstructions.Method
returnsthemethodused tocreatethe shrinkwrap.Thetypesof shrinkwrap
methodsare:
• SWCREATE_SURF_SUBSET—SurfaceSubset
• SWCREATE_FACETED_SOLID—FacetedSolid
• SWCREATE_MERGED_SOLID—MergedSolid
ThepropertyIpfcShrinkwrapModelExportInstructions.Quality
specifiesthequality levelfor thesystem tousewhen identifyingsurfacesor
componentsthatcontributeto theshrinkwrapmodel. Qualityrangesfrom 1which
producesthe coarsestrepresentationofthe modelin thefastest time,to 10which
producesthe mostexactrepresentation.The defaultvalueis 1.
Theproperty
IpfcShrinkwrapModelExportInstructions.AutoHoleFilling
setsaflagthatforcesCreo Parametricto identifyall holesandsurfacesthat
intersectasinglesurface andfillsthoseholesduring shrinkwrap.Thedefault
valueistrue.
Theproperty
IpfcShrinkwrapModelExportInstructions.IgnoreSkeleton
determineswhetherthe skeletonmodelgeometrymustbe includedin the
shrinkwrapmodel.
Theproperty
IpfcShrinkwrapModelExportInstructions.IgnoreQuiltsand
determineswhetherexternalquiltsmustbe includedin theshrinkwrapmodel.
Theproperty
IpfcShrinkwrapModelExportInstructions.AssignMassProper
tiesassignsmass propertiesto theshrinkwrapmodel. Thedefault valueisfalse
andthemasspropertiesof theoriginal modelisassignedto theshrinkwrapmodel.
Ifthevalueis settotrue, theusermustassignavalue forthe massproperties.
Interface 295

Theproperty
IpfcShrinkwrapModelExportInstructions.IgnoreSmallSurfa
cessetsaflag thatforcesCreoParametricto skipsurfaces smallerthanacertain
size.Thedefault valueisfalse. Thesize ofthe surfaceisspecifiedas apercentage
ofthemodel’ssize. Thissize canbe modifiedusing theproperty
IpfcShrinkwrapModelExportInstructions.SmallSurfPercent
age.
Theproperty
IpfcShrinkwrapModelExportInstructions.DatumReferences
specifiesandselectsthedatumplanes,points,curves,axes,andcoordinatesystem
referencesto be includedinthe shrinkwrapmodel.
Surface Subset Options
MethodsandPropertiesIntroduced:
• CCpfcShrinkwrapSurfaceSubsetInstructions.Create()
• IpfcShrinkwrapSurfaceSubsetInstructions.AdditionalSurfaces
• IpfcShrinkwrapSurfaceSubsetInstructions.OutputModel
Thestaticmethod
CCpfcShrinkwrapSurfaceSubsetInstructions.Create() returns
anobjectusedto createashrinkwrap modelofsurface subsettype. Specifythe
nameofthe outputmodelin whichtheshrinkwrap isto be createdasan inputto
thismethod.
Theproperty
IpfcShrinkwrapSurfaceSubsetInstructions.AdditionalSurfa
cesselectsindividualsurfacesto be includedin theshrinkwrap model.
Theproperty
IpfcShrinkwrapSurfaceSubsetInstructions.OutputModel sets
thetemplatemodel.
Faceted Solid Options
TheIpfcShrinkwrapFacetedFormatInstructions interfaceconsists
ofthefollowingtypes:
• SWFACETED_PART—CreoParametricpart withnormalgeometry. Thisis the
defaultformattype.
• SWFACETED_STL—AnSTL file.
• SWFACETED_VRML—AVRMLfile.
Usethe Createmethod tocreatethe objectofthespecifiedtype. Upcastthe
objectto usethegeneral methodsavailableinthis interface.
296 Creo®ParametricVBAPIUser’sGuide

PropertiesIntroduced:
• IpfcShrinkwrapFacetedFormatInstructions.Format
• IpfcShrinkwrapFacetedFormatInstructions.FramesFile
ThepropertyIpfcShrinkwrapFacetedFormatInstructions.Format
| returnstheoutput | fileformatof |     | theshrinkwrap |     | model. |
| ---------------- | ------------ | --- | ------------- | --- | ------ |
Theproperty
IpfcShrinkwrapFacetedFormatInstructions.FramesFile
enablesyoutoselectaframe fileto createafaceted solidmotionenvelopemodel
thatrepresents thefull motionof themechanismcaptured inthe framefile.
| Specifythe | nameandcomplete |     | pathoftheframe |     | file. |
| ---------- | --------------- | --- | -------------- | --- | ----- |
| Faceted    | Part Options    |     |                |     |       |
MethodsandPropertiesIntroduced:
• CCpfcShrinkwrapFacetedPartInstructions.Create()
• IpfcShrinkwrapFacetedPartInstructions.Lightweight
Thestaticmethod
CCpfcShrinkwrapFacetedPartInstructions.Create() returnsan
objectused tocreatea shrinkwrapmodelofshrinkwrap facetedtype. Theinput
| parametersof | thismethodare: |     |     |     |     |
| ------------ | -------------- | --- | --- | --- | --- |
• OutputModel—Specifytheoutput modelwheretheshrinkwrap mustbe
created.
• Lightweight—Specifythis valueasTrueiftheshrinkwrap modelisa
| LightweightCreo |     | Parametricpart. |     |     |     |
| --------------- | --- | --------------- | --- | --- | --- |
ThepropertyIpfcShrinkwrapFacetedPartInstructions.Light
eightspecifiesif theCreo Parametricpart isexportedas alightweight faceted
geometry.
| VRML | Export Options |     |     |     |     |
| ---- | -------------- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcShrinkwrapVRMLInstructions.Create()
• IpfcShrinkwrapVRMLInstructions.OutputFile
ThestaticmethodCCpfcShrinkwrapVRMLInstructions.Create()
returnsan objectusedto createa shrinkwrapmodelofshrinkwrap VRMLformat.
| Specifythe | nameofthe | outputmodelas |     | aninput | tothis method. |
| ---------- | --------- | ------------- | --- | ------- | -------------- |
ThepropertyIpfcShrinkwrapVRMLInstructions.OutputFile
| specifiesthename | oftheoutput |     | fileto | becreated. |     |
| ---------------- | ----------- | --- | ------ | ---------- | --- |
Interface 297

| STL Export | Options |     |     |     |     |     |     |
| ---------- | ------- | --- | --- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcShrinkwrapVRMLInstructions.Create()
• IpfcShrinkwrapVRMLInstructions.OutputFile
ThestaticmethodCCpfcShrinkwrapVRMLInstructions.Create()
returnsan objectusedto createa shrinkwrapmodelofshrinkwrap STLformat.
| Specifythe | nameofthe | outputmodelas |     | aninput | tothis | method. |     |
| ---------- | --------- | ------------- | --- | ------- | ------ | ------- | --- |
ThepropertyIpfcShrinkwrapSTLInstructions.OutputFile
| specifiesthename | oftheoutput |         | fileto | becreated. |     |     |     |
| ---------------- | ----------- | ------- | ------ | ---------- | --- | --- | --- |
| Merged           | Solid       | Options |        |            |     |     |     |
MethodsandPropertiesIntroduced:
• CCpfcShrinkwrapMergedSolidInstructions.Create()
• IpfcShrinkwrapMergedSolidInstructions.AdditionalComponents
Thestaticmethod
| CCpfcShrinkwrapMergedSolidInstructions.Create() |     |     |     |     |     |     | returnsan |
| ----------------------------------------------- | --- | --- | --- | --- | --- | --- | --------- |
objectused tocreatea shrinkwrapmodelofmerged solidsformat. Specifythe
| nameofthe | outputmodelasan |     | inputto | this | method. |     |     |
| --------- | --------------- | --- | ------- | ---- | ------- | --- | --- |
Theproperty
IpfcShrinkwrapMergedSolidInstructions.AdditionalCompo
nentsspecifiesindividualcomponentsofthe assemblyto be merged intothe
shrinkwrapmodel.
| VRML    | Representation |     |     |     |     |     |     |
| ------- | -------------- | --- | --- | --- | --- | --- | --- |
| Example | Code           |     |     |     |     |     |     |
Thesamplecodein thefilepfcInterfaceExamples.vblocatedat<creo_
| vbapi_loadpoint>/vbapi_examplesdemonstratesthe |     |        |               |     |                 | useof     |        |
| ---------------------------------------------- | --- | ------ | ------------- | --- | --------------- | --------- | ------ |
| IpfcDisplaymethods                             |     | with3D | displaylists. |     | Thestaticmethod |           |        |
| AddCircleDisplay()leveragesthefactthatwhena    |     |        |               |     |                 | modelwith | amodel |
programattachediserasedor deletedthe stopmethodof themodelprogramis
called.This examplecodeusesthestopmethodto produceaVRMLrepresentation
| ofthemodelin | astandarddirectoryfor |       |     | Web | publishing. |     |     |
| ------------ | --------------------- | ----- | --- | --- | ----------- | --- | --- |
| Importing    |                       | Files |     |     |             |     |     |
MethodIntroduced:
• IpfcModel.Import()
| 298 |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

ThemethodIpfcModel.Import() readsa fileintoCreo Parametric.The
formatmustbe thesameas itwouldbe if thesefileswere createdby Creo
Parametric.The parametersare:
• FilePath—Absolutepathof thefileto beimportedalong withitsextension.
• ImportData—TheIpfcImportInstructions objectthatcontrolsthe
importoperation.
Import Instructions
MethodsIntroduced:
• CCpfcRelationImportInstructions.Create()
• CCpfcIGESSectionImportInstructions.Create()
• CCpfcProgramImportInstructions.Create()
• CCpfcConfigImportInstructions.Create()
• CCpfcDWGSetupImportInstructions.Create()
• CCpfcSpoolImportInstructions.Create()
• CCpfcConnectorParamsImportInstructions.Create()
• CCpfcASSEMTreeCFGImportInstructions.Create()
• CCpfcWireListImportInstructions.Create()
• CCpfcCableParamsImportInstructions.Create()
• CCpfcSTEPImport2DInstructions.Create()
• CCpfcIGESImport2DInstructions.Create()
• CCpfcDXFImport2DInstructions.Create()
• CCpfcDWGImport2DInstructions.Create()
Themethodsdescribed in thissectioncreatean instructionsdataobjecttoimporta
fileof aspecifiedtypeintoCreo Parametric.Thedetailsare asshownin thetable
below:
Interface UsedtoImport
IpfcRelationImportInstructions Alistofrelationsandparametersinapartor
assembly.
IpfcIGESSectionImportInstructions AsectionmodelinIGESformat.
IpfcProgramImportInstructions Aprogramfileforapartorassemblythatcanbe
editedtochangethemodel.
IpfcConfigImportInstructions Configurationinstructions.
IpfcDWGSetupImportInstructions Adrawings/ufile.
IpfcSpoolImportInstructions Spoolinstructions.
IpfcConnectorParamsImportInstruc Connectorparameterinstructions.
tions
IpfcASSEMTreeCFGImportInstructions AssemblytreeCFGinstructions.
Interface 299

Interface UsedtoImport
IpfcWireListImportInstructions Wirelistinstructions.
IpfcCableParamsImportInstructions Cableparametersfromanassembly.
IpfcSTEPImport2DInstructions ApartorassemblyinSTEPformat.
IpfcIGESImport2DInstructions ApartorassemblyinIGESformat.
IpfcDXFImport2DInstructions AdrawinginDXFformat.
IpfcDWGImport2DInstructions AdrawinginDWGformat.
Note
• Themethod IpfcModel.Import()doesnot supportimportingof CADAM
typeoffiles.
• Ifamodelorthe filetypeSTEP, IGES, DWX,orSETalready exists,the
importedmodelisappended tothe currentmodel.For moreinformationon
methodsthatreturnmodelsofthe typesSTEP,IGES, DWX,andSET,refer to
GettingaModel Objecton page88.
Importing 2D Models
MethodIntroduced:
• IpfcaseSession.Import2DModel()
ThemethodIpfcBaseSession.Import2DModel() importsatwo
dimensionalmodelbasedonthe followingparameters:
• NewModelName—Specifiesthename ofthenewmodel.
• Type—Specifiesthetypeofthe model.The typecanbe oneofthe following:
○ STEP
○ IGES
○ DXF
○ DWG
○ SET
• FilePath—Specifiesthe locationofthefileto be importedalongwiththe file
extension
• Instructions—SpecifiestheIpfcImport2DInstructions objectthat
controlstheimportoperation.
TheinterfaceIpfcImport2DInstructions containsthefollowing
attributes:
300 Creo®ParametricVBAPIUser’sGuide

|     | ○ Import2DViews—Defines |     |     | whetherto | import2D |     | drawingviews. |
| --- | ----------------------- | --- | --- | --------- | -------- | --- | ------------- |
○ ScaleToFit—If thecurrentmodelhasadifferentsheetsize thanthat
specifiedbythe importedfile,setthe parameterto trueto retainthecurrent
sheetsize.Setthe parameterto falseto retainthe sheetsizeof theimported
file.
○ FitToLeftCorner—Ifthisparameterissetto true,thebottomleftcornerof
theimportedfileisadjusted to thebottomleftcornerof thecurrentmodel.
|     | Ifitissetto | false, | thesize | ofimportedfileisretained. |     |     |     |
| --- | ----------- | ------ | ------- | ------------------------- | --- | --- | --- |
Note
|           | The methodIpfcBaseSession.Import2DModel() |                  |          |     |              |     | doesnot |
| --------- | ----------------------------------------- | ---------------- | -------- | --- | ------------ | --- | ------- |
|           | support                                   | importingofCADAM |          |     | typeoffiles. |     |         |
| Importing |                                           | 3D               | Geometry |     |              |     |         |
MethodsIntroduced:
• IpfcBaseSession.GetImportSourceType()
• IpfcBaseSession.ImportNewModel()
| Forsomeinputformats,                         |     |     | themethod |     |     |     |                |
| -------------------------------------------- | --- | --- | --------- | --- | --- | --- | -------------- |
| IpfcBaseSession.GetImportSourceType()returns |     |     |           |     |     |     | thetypeofmodel |
thatcan be importedusinga designatedfile. Theinputparametersofthis method
are:
• FileToImport—Specifiesthepathof thefilealongwith itsnameand
extension.
| •                                         | NewModelImportType—Specifies |     |     |     | thetypeofmodelto |     | beimported.        |
| ----------------------------------------- | ---------------------------- | --- | --- | --- | ---------------- | --- | ------------------ |
| ThemethodIpfcBaseSession.ImportNewModel() |                              |     |     |     |                  |     | is usedto importan |
external3D formatfileandcreatesa newmodelorsetofmodelsof type
| IpfcModel.Theinput |     |     | parametersof |     | thismethodare: |     |     |
| ------------------ | --- | --- | ------------ | --- | -------------- | --- | --- |
• FileToImport—Specifiesthepathto thefilealongwith itsnameandextension
• EpfcNewModelImportType—Specifiesthe typeof modeltobe imported.
|     | Thetypesof | modelsthatcan |     | beimportedare |     | asfollows: |     |
| --- | ---------- | ------------- | --- | ------------- | --- | ---------- | --- |
○ EpfcIMPORT_NEW_IGES
○ EpfcIMPORT_NEW_VDA
○ EpfcIMPORT_NEW_NEUTRAL
○ EpfcIMPORT_NEW_CADDS
Interface 301

○ EpfcIMPORT_NEW_STEP
○ EpfcIMPORT_NEW_STL
○ EpfcIMPORT_NEW_VRML
○ EpfcIMPORT_NEW_POLTXT
○ EpfcIMPORT_NEW_CATIA_SESSION
○ EpfcIMPORT_NEW_CATIA_MODEL
○ EpfcIMPORT_NEW_DXF
○ EpfcIMPORT_NEW_ACIS
○ EpfcIMPORT_NEW_PARASOLID
○ EpfcIMPORT_NEW_ICEM
○ EpfcIMPORT_NEW_DESKTOP
○ EpfcIMPORT_NEW_CATIA_PART
○ EpfcIMPORT_NEW_CATIA_PRODUCT
○ EpfcIMPORT_NEW_UG
○ EpfcIMPORT_NEW_PRODUCTVIEW
○ EpfcIMPORT_NEW_CATIA_CGR
○ EpfcIMPORT_NEW_JT
○ EpfcIMPORT_NEW_SW_PART
○ EpfcIMPORT_NEW_SW_ASSEM
○ EpfcIMPORT_NEW_INVENTOR_PART
○ EpfcpfcIMPORT_NEW_INVENTOR_ASSEM
• EpfcModelType—Specifiesthe typeofthemodel. Itcan bea part,
assemblyor drawing.
• NewModelName—Specifiesaname forthe importedmodel.
• IpfcLayerImportFilter—Specifies thelayerfilter. Thisparameteris
optional.
Modifying the Imported Layers
Layershelpyouorganizemodelitems sothatyoucanperformoperationson those
itemscollectively.Theseoperationsprimarily includewaysof showingtheitems
inthe model,suchas displayingorblanking, selecting,andsuppressing.The
methodsdescribed inthis sectionmodifytheattributesof theimportedlayers.
MethodsandPropertiesIntroduced:
• IpfcImportedLayer.Name
302 Creo®ParametricVBAPIUser’sGuide

• IpfcImportedLayer.SetNewName()
• IpfcImportedLayer.SurfaceCount
• IpfcImportedLayer.CurveCount
• IpfcImportedLayer.TrimmedSurfaceCount
• IpfcImportedLayer.SetAction()
| Layersareidentifiedbytheir |     |     | names.Theproperty |     |     |     |
| -------------------------- | --- | --- | ----------------- | --- | --- | --- |
IpfcImportedLayer.Namereturns thenameof thelayerwhilethe method
| IpfcImportedLayer.SetNewName() |         |           |                 | canbe usedto | setthenameof | the |
| ------------------------------ | ------- | --------- | --------------- | ------------ | ------------ | --- |
| layer.The                      | namecan | benumeric | oralphanumeric. |              |              |     |
ThepropertyIpfcImportedLayer.SurfaceCount returns thenumberof
| curvesonthe                                      | layer.            |     |                       |                     |     |            |
| ------------------------------------------------ | ----------------- | --- | --------------------- | ------------------- | --- | ---------- |
| ThepropertyIpfcImportedLayer.TrimmedSurfaceCount |                   |     |                       |                     |     | returnsthe |
| numberof                                         | trimmedsurfaceson |     | thelayer andthemethod |                     |     |            |
| IpfcImportedLayer.CurveCountreturns              |                   |     |                       | thenumberofcurveson |     | the        |
layer.
| ThemethodIpfcImportedLayer.SetAction() |     |     |     |     | setsthedisplayofthe |     |
| -------------------------------------- | --- | --- | --- | --- | ------------------- | --- |
importedlayers.Theinputparameter forthis methodis ImportAction.Thetypes
| ofactionsthatcanbe                          |     | performedonthe | importedlayersare: |                |            |     |
| ------------------------------------------- | --- | -------------- | ------------------ | -------------- | ---------- | --- |
| • IMPORT_LAYER_DISPLAY—Displaysthe          |     |                |                    | importedlayer. |            |     |
| • IMPORT_LAYER_SKIP—Doesnotimportentitieson |     |                |                    |                | thislayer. |     |
• IMPORT_LAYER_BLANK—Blankstheselectedlayer.
• IMPORT_LAYER_IGNORE—Imports onlyentitieson thislayerbut notthe
layer.
ThedefaultactiontypeisIMPORT_LAYER_DISPLAY.
| Plotting                               | Files |     |     |                      |     |     |
| -------------------------------------- | ----- | --- | --- | -------------------- | --- | --- |
| FromPro/ENGINEERWildfire5.0onwards,the |       |     |     | IpfcPlotInstructions |     |     |
objectcontainingthe instructionsfor plottingfileshasbeendeprecated. Allthe
methodslistedbelowfor creatingandaccessing theinstructionattributesin
IpfcPlotInstructionshavealsobeendeprecated. Usethe newinterface
| typeIpfcPrinterInstructions |     |     | anditsmethodsdescribedinthe |     |     | next |
| --------------------------- | --- | --- | --------------------------- | --- | --- | ---- |
section.
MethodsandPropertiesDeprecated:
• CCpfcPlotInstructions.Create()
• IpfcPlotInstructions.PlotterName
• IpfcPlotInstructions.OutputQuality
• IpfcPlotInstructions.UserScale
| Interface |     |     |     |     |     | 303 |
| --------- | --- | --- | --- | --- | --- | --- |

• IpfcPlotInstructions.PenSlew
• IpfcPlotInstructions.PenVelocityX
• IpfcPlotInstructions.PenVelocityY
• IpfcPlotInstructions.SegmentedOutput
• IpfcPlotInstructions.LabelPlot
• IpfcPlotInstructions.SeparatePlotFiles
• IpfcPlotInstructions.PaperSize
• IpfcPlotInstructions.PageRangeChoice
• IpfcPlotInstructions.PaperSizeX
• IpfcPlotInstructions.FirstPage
• IpfcPlotInstructions.LastPage
Printing Files
Theprinterinstructionsforprinting afilearedefinedin
IpfcPrinterInstructionsdataobject.
MethodsandPropertiesIntroduced:
• CCpfcPrinterInstructions.Create()
• IpfcPrinterInstructions.PrinterOption
• IpfcPrinterInstructions.PlacementOption
• IpfcPrinterInstructions.ModelOption
• pfcExport.IpfcPrinterInstructions.WindowId
ThemethodCCpfcPrinterInstructions.Create() createsanew
instanceof theIpfcPrinterInstructions object.Theobjectcontains the
followinginstructionattributes:
• PrinterOption—Specifiesthe printersettings forprinting afileinterms ofthe
IpfcPrintPrinterOptionobject.Setthis attributeusing theproperty
IpfcPrinterInstructions.PrinterOption.
• PlacementOption—Specifiestheplacement optionsfor printingpurpose in
termsofthe IpfcPrintMdlOptionobject.Setthisattribute using the
propertyIpfcPrinterInstructions.PlacementOption.
• ModelOption—Specifiesthemodeloptions forprinting purposein termsof
theIpfcPrintPlacementOption object.Setthisattribute usingthe
propertyIpfcPrinterInstructions.ModelOption.
• WindowId—Specifiesthe currentwindow identifier.Setthis attributeusing the
propertyIpfcPrinterInstructions.WindowId.
304 Creo®ParametricVBAPIUser’sGuide

Printer Options
Theprintersettingsfor printingafilearedefinedin the
IpfcPrintPrinterOptionobject.
MethodsandPropertiesIntroduced:
• CCpfcPrintPrinterOption.Create()
• IpfcBaseSession.GetPrintPrinterOptions()
• IpfcPrintPrinterOption.DeleteAfter
• IpfcPrintPrinterOption.FileName
• IpfcPrintPrinterOption.PaperSize
• CCpfcPrintSize.Create()
• IpfcPrintSize.Height
• IpfcPrintSize.Width
• IpfcPrintSize.PaperSize
• IpfcPrintPrinterOption.PenTable
• IpfcPrintPrinterOption.PrintCommand
• IpfcPrintPrinterOption.PrinterType
• IpfcPrintPrinterOption.Quantity
• IpfcPrintPrinterOption.RollMedia
• IpfcPrintPrinterOption.RotatePlot
• IpfcPrintPrinterOption.SaveMethod
• IpfcPrintPrinterOption.SaveToFile
• IpfcPrintPrinterOption.SendToPrinter
• IpfcPrintPrinterOption.Slew
• IpfcPrintPrinterOption.SwHandshake
• IpfcPrintPrinterOption.UseTtf
ThemethodCCpfcPrintPrinterOption.Create() createsa new
instanceof theIpfcPrintPrinterOption object.
ThemethodIpfcBaseSession.GetPrintPrinterOptions() retrieves
theprintersettings.
TheIpfcPrintPrinterOption objectcontainsthe followingoptions:
• DeleteAfter—Determinesif thefileisdeleted afterprinting.Setitto trueto
deletethefileafterprinting.Use theproperty
IpfcPrintPrinterOption.DeleteAfter to assignthisoption.
Interface 305

• FileName—Specifiesthename ofthefileto be printed.Use theproperty
| IpfcPrintPrinterOption.FileNameto |     |     |     |     |     | setthename. |     |     |     |
| --------------------------------- | --- | --- | --- | --- | --- | ----------- | --- | --- | --- |
Note
|     | Ifthe methodIpfcModel.Export()   |                                    |     |     | iscalledfor    |     | IpfcExportType |            |     |
| --- | -------------------------------- | ---------------------------------- | --- | --- | -------------- | --- | -------------- | ---------- | --- |
|     | object,thentheargumentFileNameis |                                    |     |     | ignored,andcan |     |                | bepassedas |     |
|     | NULL. You                        | mustusethemethodIpfcModel.Export() |     |     |                |     |                | tosetthe   |     |
FileName.
• PaperSize—Specifiesthe parametersof thepaperto be printedinterms ofthe
| IpfcPrintSize                    |     |     | object.The | property |     |                            |     |     |     |
| -------------------------------- | --- | --- | ---------- | -------- | --- | -------------------------- | --- | --- | --- |
| IpfcPrintPrinterOption.PaperSize |     |     |            |          |     | assignsthePaperSizeoption. |     |     |     |
Usethe methodCCpfcPrintSize.Create() to createanewinstance of
theIpfcPrintSize object.Thisobjectcontains thefollowingoptions:
| ○   | Height—Specifies     |     | theheightofpaper. |                      | Usethe          | property |     |     |     |
| --- | -------------------- | --- | ----------------- | -------------------- | --------------- | -------- | --- | --- | --- |
|     | IpfcPrintSize.Height |     |                   | tosetthe             | paperheight.    |          |     |     |     |
| ○   | Width—Specifiesthe   |     | widthofpaper.     |                      | Use theproperty |          |     |     |     |
|     | IpfcPrintSize.Width  |     |                   | to setthepaperwidth. |                 |          |     |     |     |
○ PaperSize—Specifies thesize ofthepaperused fortheplotin termsof the
|     | EpfcPlotPaperSize         |     |     | object.Usethe | property         |     |     |     |     |
| --- | ------------------------- | --- | --- | ------------- | ---------------- | --- | --- | --- | --- |
|     | IpfcPrintSize.PaperSizeto |     |     |               | setthepapersize. |     |     |     |     |
Note
|     | If youwantto                     |     | plotalayoutwithoutadding |     |               | aborderon |     | thepaper,use |     |
| --- | -------------------------------- | --- | ------------------------ | --- | ------------- | --------- | --- | ------------ | --- |
|     | the followingpapersizesdefinedin |     |                          |     | theenumerated |           |     | datatype     |     |
EpfcPlotPaperSize:
|     | ◆ EpfcCEEMPTYPLOT—Thepapersize    |     |     |     |     | is22.5 | x36        | in  |     |
| --- | --------------------------------- | --- | --- | --- | --- | ------ | ---------- | --- | --- |
|     | ◆ EpfcCEEMPTYPLOT_MM—Thepapersize |     |     |     |     |        | is625x1000 |     | mm  |
• PenTable—Specifiesthefilecontainingthepentable. Usethe property
| IpfcPrintPrinterOption.PenTableto |     |     |     |     |     | setthisoption. |     |     |     |
| --------------------------------- | --- | --- | --- | --- | --- | -------------- | --- | --- | --- |
• PrintCommand—Specifiesthecommandto be usedfor printing.Use the
propertyIpfcPrintPrinterOption.PrintCommand
tosetthe
command.
| • PrinterType—Specifiestheprintertype. |     |     |     |     | Usethe | property                        |                |     |     |
| -------------------------------------- | --- | --- | --- | --- | ------ | ------------------------------- | -------------- | --- | --- |
| IpfcPrintPrinterOption.PrinterType     |     |     |     |     |        | to                              | assignthetype. |     |     |
| 306                                    |     |     |     |     |        | Creo®ParametricVBAPIUser’sGuide |                |     |     |

• Quantity—Specifiesthe numberofcopiesto be printed.Usethe property
| IpfcPrintPrinterOption.Quantityto |     |     |     |     | assignthequantity. |     |     |
| --------------------------------- | --- | --- | --- | --- | ------------------ | --- | --- |
• RollMedia—Determinesifroll mediaisto beused forprinting. Setittotrue to
| userollmedia.Usethe              |     | property |     |     |                      |     |     |
| -------------------------------- | --- | -------- | --- | --- | -------------------- | --- | --- |
| IpfcPrintPrinterOption.RollMedia |     |          |     |     | to assignthisoption. |     |     |
• RotatePlot—Determinesif theplotisrotated by90 degrees.Setitto trueto
| rotatethe                            | plot.Use | theproperty |                     |         |               |                |         |
| ------------------------------------ | -------- | ----------- | ------------------- | ------- | ------------- | -------------- | ------- |
| IpfcPrintPrinterOption.SetRotatePlot |          |             |                     |         | to            | setthisoption. |         |
| • SaveMethod—Specifiesthesavemethod  |          |             |                     | interms | ofthe         |                |         |
| EpfcPrintSaveMethod                  |          |             | enumeratedclass.Use |         | theproperty   |                |         |
| IpfcPrintPrinterOption.SaveMethodto  |          |             |                     |         | specify       | thesave        | method. |
| Theavailablemethodsareas             |          |             | follows:            |         |               |                |         |
| ○ EpfcPRINT_SAVE_SINGLE_FILE—Plot    |          |             |                     |         | issavedto     | asinglefile.   |         |
| ○ EpfcPRINT_SAVE_MULTIPLE_FILE—Plot  |          |             |                     |         | issavedto     | multiplefiles. |         |
| ○ EpfcPRINT_SAVE_APPEND_TO_FILE—Plot |          |             |                     |         | is appendedto |                | afile.  |
• SaveToFile—Determinesif thefileissavedafterprinting.Setitto true tosave
| thefileafterprinting.Use            |     |     | theproperty |     |                   |     |     |
| ----------------------------------- | --- | --- | ----------- | --- | ----------------- | --- | --- |
| IpfcPrintPrinterOption.SaveToFileto |     |     |             |     | assignthisoption. |     |     |
• SendToPrinter—Determines iftheplotis directlysent to theprinter.Set itto
| trueto sendthe                       |     | plotto theprinter.Use |     | theproperty |     |                |     |
| ------------------------------------ | --- | --------------------- | --- | ----------- | --- | -------------- | --- |
| IpfcPrintPrinterOption.SendToPrinter |     |                       |     |             | to  | setthisoption. |     |
• Slew—Specifiesthespeedofthe penincentimeterspersecondin XandY
direction.Usethe propertyIpfcPrintPrinterOption.Slewto setthis
option.
• SwHandshake—Determinesif thesoftware handshakemethodisto be used
forprinting. Setittotrue to usethe software handshakemethod.Use the
| propertyIpfcPrintPrinterOption.SwHandshake |     |     |     |     |     | to setthisoption. |     |
| ------------------------------------------ | --- | --- | --- | --- | --- | ----------------- | --- |
• UseTtf—SpecifieswhetherTrueTypefontsor strokedtextisused for
printing.Setthis optionto trueto useTrueTypefonts andto falseto stroke
| alltext. Use | thepropertyIpfcPrintPrinterOption.UseTtf |     |     |     |     |     | toset |
| ------------ | ---------------------------------------- | --- | --- | --- | --- | --- | ----- |
thisoption.
| Placement           | Options |                        |     |              |     |     |     |
| ------------------- | ------- | ---------------------- | --- | ------------ | --- | --- | --- |
| Theplacementoptions |         | for printingpurposeare |     | definedinthe |     |     |     |
IpfcPrintPlacementOptionobject.
MethodsandPropertiesIntroduced:
• CCpfcPrintPlacementOption.Create()
| Interface |     |     |     |     |     |     | 307 |
| --------- | --- | --- | --- | --- | --- | --- | --- |

• IpfcBaseSession.GetPrintPlacementOptions()
• IpfcPrintPlacementOption.BottomOffset
• IpfcPrintPlacementOption.ClipPlot
• IpfcPrintPlacementOption.KeepPanzoom
• IpfcPrintPlacementOption.LabelHeight
• IpfcPrintPlacementOption.PlaceLabel
• IpfcPrintPlacementOption.Scale
• IpfcPrintPlacementOption.ShiftAllCorner
• IpfcPrintPlacementOption.SideOffset
• IpfcPrintPlacementOption.X1ClipPosition
• IpfcPrintPlacementOption.X2ClipPosition
• IpfcPrintPlacementOption.Y1ClipPosition
• IpfcPrintPlacementOption.Y2ClipPosition
ThemethodCCpfcPrintPlacementOption.Create() createsanew
instanceof theIpfcPrintPlacementOption object.
ThemethodIpfcBaseSession.GetPrintPlacementOptions()
retrievestheplacement options.
TheIpfcPrintPlacementOption objectcontainsthe followingoptions:
• BottomOffset—Specifiesthe offsetfrom thelower-leftcornerof theplot. Use
thepropertyIpfcPrintPlacementOption.BottomOffset to setthis
option.
• ClipPlot—Specifieswhethertheplotisclipped. Setthisoption to trueto clip
theplotorto falseto avoidclippingofplot.Use theproperty
IpfcPrintPlacementOption.ClipPlotto setthis option.
• KeepPanzoom—Determineswhetherpan andzoomvalues ofthewindoware
used.Setthisoption to trueusepanandzoom andfalsetoskip them.Usethe
propertyIpfcPrintPlacementOption.KeepPanzoom to setthis
option.
• LabelHeight—Specifiestheheightofthe labelin inches.Use theproperty
IpfcPrintPlacementOption.LabelHeight to setthisoption.
• PlaceLabel—Specifieswhetheryouwantto place thelabelon theplot. Use
thepropertyIpfcPrintPlacementOption.PlaceLabel to setthis
option.
• Scale—Specifiesthescale usedforthe plot.Use theproperty
IpfcPrintPlacementOption.Scale to setthisoption.
308 Creo®ParametricVBAPIUser’sGuide

• ShiftAllCorner—Determineswhetherall cornersareshifted.Set thisoption to
trueto shift allcornersor to falseto skipshiftingof corners.Use theproperty
IpfcPrintPlacementOption.ShiftAllCorner tosetthis option.
• SideOffset—Specifiestheoffsetfrom thesides.Use theproperty
IpfcPrintPlacementOption.SideOffset tosetthis option.
• X1ClipPosition—Specifiesthefirst Xparameterfor definingthe clipposition.
Usethe propertyIpfcPrintPlacementOption.X1ClipPosition to
setthisoption.
• X2ClipPosition—SpecifiesthesecondXparameterfor definingthe clip
position.Usethe property
IpfcPrintPlacementOption.X2ClipPosition tosetthis option.
• Y1ClipPosition—Specifiesthe firstYparameterfor definingthe clipposition.
Usethe propertyIpfcPrintPlacementOption.Y1ClipPosition to
setthisoption.
• Y2ClipPosition—Specifiesthe secondYparameter for definingtheclip
position.Usethe property
IpfcPrintPlacementOption.Y2ClipPosition tosetthis option.
Model Options
Themodeloptions forprinting purposearedefinedin the
IpfcPrintMdlOptionobject.
MethodsandPropertiesIntroduced:
• CCpfcPrintMdlOption.Create()
• IpfcBaseSession.GetPrintMdlOptions()
• IpfcPrintMdlOption.DrawFormat
• IpfcPrintMdlOption.FirstPage
• IpfcPrintMdlOption.LastPage
• IpfcPrintMdlOption.LayerName
• IpfcPrintMdlOption.LayerOnly
• IpfcPrintMdlOption.Mdl
• IpfcPrintMdlOption.Quality
• IpfcPrintMdlOption.Segmented
• IpfcPrintMdlOption.Sheets
• IpfcPrintMdlOption.UseDrawingSize
• IpfcPrintMdlOption.UseSolidScale
Interface 309

| ThemethodCCpfcPrintMdlOption.Create() |     |     |     |     |     | createsa | newinstanceof |
| ------------------------------------- | --- | --- | --- | --- | --- | -------- | ------------- |
theIpfcPrintMdlOptionobject.
| ThemethodIpfcBaseSession.GetPrintMdlOptions() |     |     |     |     |     |     | retrievesthe |
| --------------------------------------------- | --- | --- | --- | --- | --- | --- | ------------ |
modeloptions.
| TheIpfcPrintMdlOption |     |     | objectcontainsthefollowingoptions: |     |     |     |     |
| --------------------- | --- | --- | ---------------------------------- | --- | --- | --- | --- |
• DrawFormat—Displays thedrawing formatused forprinting.Use the
propertyIpfcPrintMdlOption.DrawFormatto
setthisoption.
| • FirstPage—Specifiesthefirst |     |     | pagenumber.    |           | Use theproperty |          |     |
| ----------------------------- | --- | --- | -------------- | --------- | --------------- | -------- | --- |
| IpfcPrintMdlOption.FirstPage  |     |     |                | tosetthis |                 | option.  |     |
| • LastPage—Specifiesthelast   |     |     | pagenumber.Use |           | theproperty     |          |     |
| IpfcPrintMdlOption.LastPageto |     |     |                |           | setthisoption.  |          |     |
| • LayerName—Specifiesthe      |     |     | nameofthe      | layer.    | Usethe          | property |     |
| IpfcPrintMdlOption.LayerName  |     |     |                | tosetthe  |                 | name.    |     |
• LayerOnly—Printsthe specifiedlayeronly.Setthis optionto trueto print
thespecifiedlayer. Usethe propertyIpfcPrintMdlOption.LayerOnly
| tosetthis option.         |     |     |                |         |             |     |     |
| ------------------------- | --- | --- | -------------- | ------- | ----------- | --- | --- |
| • Mdl—Specifiesthemodelto |     |     | be printed.Use |         | theproperty |     |     |
| IpfcPrintMdlOption.Mdlto  |     |     |                | setthis | option.     |     |     |
• Quality—Determinesthequality ofthemodelto be printed.It checksfor no
line,no overlap,simpleoverlap,andcomplexoverlap.Use theproperty
IpfcPrintMdlOption.Quality
|     |     |     |     | tosetthis | option. |     |     |
| --- | --- | --- | --- | --------- | ------- | --- | --- |
• Segmented—Ifsetto true, theprinterprintsthedrawing in fullsize,but in
segmentsthatarecompatiblewith theselectedpapersize.This optionis
| availableonly                | if youareplottingasinglepage. |           |             |           | Use            | theproperty |       |
| ---------------------------- | ----------------------------- | --------- | ----------- | --------- | -------------- | ----------- | ----- |
| IpfcPrintMdlOption.Segmented |                               |           |             | tosetthis |                | option.     |       |
| • Sheets—Specifiesthesheets  |                               |           | thatneedto  | be        | printedinterms |             | ofthe |
| EpfcPrintSheets              |                               | class.Use | theproperty |           |                |             |       |
IpfcPrintMdlOption.Sheetsto specifythe sheets.Thesheetscan be
ofthefollowingtypes:
| ○ EpfcPRINT_CURRENT_SHEET—Only     |     |     |     |              | thecurrentsheetisprinted. |                       |     |
| ---------------------------------- | --- | --- | --- | ------------ | ------------------------- | --------------------- | --- |
| ○ EpfcPRINT_ALL_SHEETS—All         |     |     |     | thesheetsare |                           | printed.              |     |
| ○ EpfcPRINT_SELECTED_SHEETS—Sheets |     |     |     |              |                           | ofa specifiedrangeare |     |
printed.
• UseDrawingSize—Overridesthe papersizespecifiedin theprinteroptions
withthe drawingsize. Setthisoption to truetousethedrawing size.Use the
| propertyIpfcPrintMdlOption.UseDrawingSize |     |     |     |     |                                 |     | tosetthis option. |
| ----------------------------------------- | --- | --- | --- | --- | ------------------------------- | --- | ----------------- |
| 310                                       |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |                   |

• UseSolidScale—Printswiththe scaleusedin thesolidmodel. Setthisoption
| totrueto                         | usesolidscale.Usethe |            |         | property     |                   |     |
| -------------------------------- | -------------------- | ---------- | ------- | ------------ | ----------------- | --- |
| IpfcPrintMdlOption.UseSolidScale |                      |            |         |              | to setthisoption. |     |
| Plotter                          | Configuration        |            | File    | (PCF)        | Options           |     |
| Theprinting                      | optionsfor           | PCFfileare |         | definedinthe |                   |     |
| IpfcPrinterPCFOptions            |                      |            | object. |              |                   |     |
MethodsandPropertiesIntroduced:
• CCpfcPrinterPCFOptions.Create()
• IpfcPrinterPCFOptions.PrinterOption
• IpfcPrinterPCFOptions.PlacementOption
• IpfcPrinterPCFOptions.ModelOption
ThemethodCCpfcPrinterPCFOptions.Create() createsanewinstance
| oftheIpfcPrinterPCFOptions |     |     |     | object.        |                      |     |
| -------------------------- | --- | --- | --- | -------------- | -------------------- | --- |
| TheIpfcPrinterPCFOptions   |     |     |     | objectcontains | thefollowingoptions: |     |
• PrinterOption—Specifiesthe printersettings forprinting afileinterms ofthe
IpfcPrintPrinterOptionobject.Setthis
|     |     |     |     |     | attributeusing | theproperty |
| --- | --- | --- | --- | --- | -------------- | ----------- |
IpfcPrinterPCFOptions.PrinterOption.
• PlacementOption—Specifiestheplacement optionsfor printingpurpose in
termsofthe IpfcPrintMdlOptionobject.Setthisattribute using the
propertyIpfcPrinterPCFOptions.SetPlacementOption.
• ModelOption—Specifiesthemodeloptions forprinting purposein termsof
| theIpfcPrintPlacementOption |     |     |     | object.Setthisattribute |     | usingthe |
| --------------------------- | --- | --- | --- | ----------------------- | --- | -------- |
propertyIpfcPrinterPCFOptions.ModelOption.
| Solid | Operations |     |     |     |     |     |
| ----- | ---------- | --- | --- | --- | --- | --- |
MethodIntroduced:
• IpfcSolid.CreateImportFeat()
| ThemethodIpfcSolid.CreateImportFeat() |              |                                  |     |     | createsa | newimport |
| ------------------------------------- | ------------ | -------------------------------- | --- | --- | -------- | --------- |
| featurein                             | the solidand | takesthefollowinginputarguments: |     |     |          |           |
• IntfData—Thesourceofdatafrom whichto createtheimportfeature.Itis
givenby theIpfcIntfDataSourceobject.The typeof sourcedatathat
| canbe | importedisgivenby |     | theEpfcIntfType |     | classandcan | be ofthe |
| ----- | ----------------- | --- | --------------- | --- | ----------- | -------- |
followingtypes:
○ EpfcINTF_NEUTRAL
Interface 311

○ EpfcINTF_NEUTRAL_FILE
○ EpfcINTF_IGES
○ EpfcINTF_STEP
○ EpfcINTF_VDA
○ EpfcINTF_ICEM
○ EpfcINTF_ACIS
○ EpfcINTF_DXF
○ EpfcINTF_CDRS
○ EpfcINTF_STL
○ EpfcINTF_VRML
○ EpfcINTF_PARASOLID
○ EpfcINTF_AI
○ EpfcINTF_CATIA_PART
○ EpfcINTF_UG
○ EpfcINTF_PRODUCTVIEW
○ EpfcINTF_CATIA_CGR
○ EpfcINTF_JT
• CoordSys—Thepointerto areferencecoordinatesystem. Ifthisis NULL,the
functionusesthedefault coordinatesystem.
• FeatAttr—Theattributesfor creationofthenewimportfeaturegiven bythe
IpfcImportFeatAttrobject.Ifthis pointerisNULL, thefunctionuses
thedefaultattributes.
Example Code: Returning a Feature Object
Thesamplecodein thefilepfcInterfaceExamples.vblocatedat<creo_
vbapi_loadpoint>/vbapi_examplesreturna featureobjectwhen
providedwitha solidcoordinatesystem nameandan importfeature'sfilename.
Themethodwillfindthe coordinatesystem in themodel, settheImport Feature
Attributes,and createan importfeature.Thefeatureis thenreturned.
Window Operations
MethodIntroduced:
• IpfcWindow.ExportRasterImage()
ThemethodIpfcWindow.ExportRasterImage() outputs astandard Creo
Parametricrasteroutputfile.
312 Creo®ParametricVBAPIUser’sGuide

| Example | Code: Generating | Raster | Files |
| ------- | ---------------- | ------ | ----- |
Thesamplecodein thefilepfcInterfaceExamples.vblocatedat<creo_
vbapi_loadpoint>/vbapi_examplesgeneratesrasterimagefilesusinga
| specifiedwindow | andfiletype. |     |     |
| --------------- | ------------ | --- | --- |
Interface 313

24
Simplified Representations
Overview................................................................................................................315
RetrievingSimplifiedRepresentations.......................................................................316
CreatingandDeletingSimplifiedRepresentations......................................................317
ExtractingInformationAboutSimplifiedRepresentations............................................317
ModifyingSimplifiedRepresentations.......................................................................318
SimplifiedRepresentationUtilities.............................................................................320
TheVB APIgivesprogrammatic accessto allthesimplifiedrepresentation
functionalityofCreo Parametric.Createsimplifiedrepresentations either
permanentlyor onthe flyand save,retrieve,ormodifythem byadding ordeleting
items.
314 Creo®ParametricVBAPIUser’sGuide

Overview
| UsingtheVB               | API,youcan | createandmanipulateassembly           |     | simplified |
| ------------------------ | ---------- | ------------------------------------- | --- | ---------- |
| representationsjustasyou |            | canusing CreoParametricinteractively. |     |            |
Note
theVB APIsupports simplifiedrepresentationofassemblies only, notparts.
IpfcSimRepclass.This
Simplifiedrepresentationsareidentifiedbythe classis
| achildof IpfcModelItem,soyoucan |     |     | usethe methodsdealing | with |
| ------------------------------- | --- | --- | --------------------- | ---- |
IpfcModelItemsto collect,inspect,andmodifysimplifiedrepresentations.
Theinformationrequired tocreateandmodifya simplifiedrepresentationis
storedina classcalledIpfcSimpRepInstructions whichcontainsseveral
| dataobjectsand | fields,including: |                          |     |     |
| -------------- | ----------------- | ------------------------ | --- | --- |
| • String—The   | nameofthe         | simplifiedrepresentation |     |     |
• IpfcSimpRepAction—The rulethatcontrolsthedefault treatmentof
| itemsin | thesimplifiedrepresentation. |     |     |     |
| ------- | ---------------------------- | --- | --- | --- |
• IpfcSimpRepItem—Anarrayofassembly componentsandtheactions
| appliedto | themin thesimplifiedrepresentation. |     |     |     |
| --------- | ----------------------------------- | --- | --- | --- |
AIpfcSimpRepItemis identifiedbythe assemblycomponentpathtothat
item.EachIpfcSimpRepItem hasit’s ownIpfcSimpRepAction assigned
toit. IpfcSimpRepAction isa visibledataobjectthatincludes afieldof type
| IpfcSimpRepActionType.Youcan |     |     | usethemethod |     |
| ---------------------------- | --- | --- | ------------ | --- |
IpfcSimpRepAction.GetType()to settheactions.To deleteanexisting
item,youmustsettheactionasNULL.
EpfcSimpRepActionType isan enumeratedtypethatspecifiesthe possible
treatmentof itemsin asimplifiedrepresentation.The possiblevaluesareas
follows
| Values           |     |     | Action               |     |
| ---------------- | --- | --- | -------------------- | --- |
| EpfcSIMPREP_NONE |     |     | Noactionisspecified. |     |
EpfcSIMPREP_REVERSE Reversethedefaultruleforthiscomponent(for
example,includeitifthedefaultruleisexclude).
| EpfcSIMPREP_INCLUDE |     |     | Includethiscomponentinthesimplified |     |
| ------------------- | --- | --- | ----------------------------------- | --- |
representation.
| EpfcSIMPREP_EXCLUDE |     |     | Excludethiscomponentfromthesimplified |     |
| ------------------- | --- | --- | ------------------------------------- | --- |
representation.
| EpfcSIMPREP_SUBSTITUTE |     |     | Substitutethecomponentinthesimplified |     |
| ---------------------- | --- | --- | ------------------------------------- | --- |
representation.
| EpfcSIMPREP_GEOM |     |     | Useonlythegeometricalrepresentationofthe |     |
| ---------------- | --- | --- | ---------------------------------------- | --- |
component.
| EpfcSIMPREP_GRAPHICS |     |     | Useonlythegraphicsrepresentationofthe |     |
| -------------------- | --- | --- | ------------------------------------- | --- |
component.
SimplifiedRepresentations 315

| Values |     |     | Action |     |     |
| ------ | --- | --- | ------ | --- | --- |
EpfcSIMPREP_SYMB
Usethesymbolicrepresentationofthecomponent.
| EpfcSIMPREP_BOUNDBOX |     |     | Usetheboundaryboxrepresentationofthe |     |     |
| -------------------- | --- | --- | ------------------------------------ | --- | --- |
component.
| EpfcSIMPREP_DEFENV |     |     | Usethedefaultenveloperepresentationofthe |     |     |
| ------------------ | --- | --- | ---------------------------------------- | --- | --- |
component.
EpfcSIMPREP_LIGHT_GRAPH Usethelightweightgraphicsrepresentationofthe
component.
EpfcSIMPREP_AUTO Usetheautomaticrepresentationofthecomponent.
| Retrieving |     | Simplified | Representations |     |     |
| ---------- | --- | ---------- | --------------- | --- | --- |
MethodsIntroduced:
• IpfcBaseSession.RetrieveAssemSimpRep()
• IpfcBaseSession.RetrieveGeomSimpRep()
• IpfcBaseSession.RetrieveGraphicsSimpRep()
• IpfcBaseSession.RetrieveSymbolicSimpRep()
• CCpfcRetrieveExistingSimpRepInstructions.Create()
You canretrieve anamedsimplifiedrepresentationfroma modelusing the
| methodIpfcBaseSession.RetrieveAssemSimpRep(), |     |     |     | whichis |     |
| --------------------------------------------- | --- | --- | --- | ------- | --- |
analogousto theAssembly modeoptionRetrieveRep inthe SIMPLFDREPmenu.
Thismethodretrieves theobjectof anexisting simplifiedrepresentationfrom an
assemblywithoutfetching thegenericrepresentationintomemory.The method
takestwoarguments,the nameofthe assemblyandthe simplifiedrepresentation
data.
| Toretrieve                                        | an existingsimplifiedrepresentation,passaninstance |     |     | of  |     |
| ------------------------------------------------- | -------------------------------------------------- | --- | --- | --- | --- |
| CCpfcRetrieveExistingSimpRepInstructions.Create() |                                                    |     |     |     | and |
specifyitsnameas thesecond argumentto this method.Creo Parametricretrieves
thatrepresentation andanyactivesubmodelsandreturnsthe objecttothe
| simplifiedrepresentation |     | asaIpfcAssembly.Assembly |     | object. |     |
| ------------------------ | --- | ------------------------ | --- | ------- | --- |
You canretrieve geometry,graphics,andsymbolicsimplifiedrepresentationsinto
sessionusing the methodsIpfcBaseSession.RetrieveGeomSimpRep(),
IpfcBaseSession.RetrieveGraphicsSimpRep(),and
| IpfcBaseSession.RetrieveSymbolicSimpRep() |     |     |     | respectively.        | Like |
| ----------------------------------------- | --- | --- | --- | -------------------- | ---- |
| IpfcBaseSession.RetrieveAssemSimpRep(),   |     |     |     | thesemethodsretrieve |      |
thesimplifiedrepresentationwithoutbringingthe masterrepresentation into
memory. Supplythename oftheassembly whosesimplifiedrepresentationis to
beretrievedas theinputparameterforthesemethods.The methodsoutput the
| assembly. | Theydonotdisplaythe | simplifiedrepresentation. |     |     |     |
| --------- | ------------------- | ------------------------- | --- | --- | --- |
316 Creo®ParametricVBAPIUser’sGuide

| Creating | and Deleting | Simplified |     |
| -------- | ------------ | ---------- | --- |
Representations
MethodsIntroduced:
• CCpfcCreateNewSimpRepInstructions.Create()
• IpfcSolid.CreateSimpRep()
• IpfcSolid.DeleteSimpRep()
| Tocreatea                                   | simplifiedrepresentation,youmustallocateand |           | filla          |
| ------------------------------------------- | ------------------------------------------- | --------- | -------------- |
| IpfcSimpRepInstructionsobjectbycalling      |                                             | themethod |                |
| CCpfcCreateNewSimpRepInstructions.Create(). |                                             |           | Specifythename |
ofthenewsimplifiedrepresentationas aninput tothis method.You shouldalso
| setthe defaultactiontypeandaddSimpRepItemsto |     |     | theobject. |
| -------------------------------------------- | --- | --- | ---------- |
Togeneratethenewsimplifiedrepresentation,call
| IpfcSolid.CreateSimpRep().          |                       | Thismethodreturns  | theIpfcSimpRep |
| ----------------------------------- | --------------------- | ------------------ | -------------- |
| objectfor                           | thenewrepresentation. |                    |                |
| ThemethodIpfcSolid.DeleteSimpRep()  |                       | deletesasimplified |                |
| representationfromitsmodelowner.The |                       | methodrequiresonly | the            |
IpfcSimpRepobjectas input.
| Extracting | Information | About | Simplified |
| ---------- | ----------- | ----- | ---------- |
Representations
MethodsandPropertiesIntroduced:
• IpfcSimpRep.GetInstructions()
• IpfcSimpRepInstructions.DefaultAction()
• IpfcCreateNewSimpRepInstructions.NewSimpName
• IpfcSimpRepInstructions.IsTemporary
• IpfcSimpRepInstructions.Items
| Giventheobjectto | asimplifiedrepresentation, |     |     |
| ---------------- | -------------------------- | --- | --- |
IpfcSimpRep.GetInstructions()fillsoutthe
IpfcSimpRepInstructionsobject.
ThepropertiesIpfcSimpRepInstructions.GetDefaultAction,
IpfcCreateNewSimpRepInstructions.NewSimpName, and
IpfcSimpRepInstructions.IsTemporaryreturn theassociatedvalues
| containedin | theIpfcSimpRepInstructionsobject. |     |     |
| ----------- | --------------------------------- | --- | --- |
ThepropertyIpfcSimpRepInstructions.Items returnsallthe itemsthat
| makeupthe | simplifiedrepresentation. |     |     |
| --------- | ------------------------- | --- | --- |
SimplifiedRepresentations 317

| Example                                | 1:  | Working | with | Simplified | Representation  |     |
| -------------------------------------- | --- | ------- | ---- | ---------- | --------------- | --- |
| Thesamplecodein                        |     | thefile |      |            |                 |     |
| pfcSimplifiedRepresentationExamples.vb |     |         |      |            | locatedat<creo_ |     |
vbapi_loadpoint>/vbapi_examplesdemonstratesthe functionalityused
whenworking withexisting simplifiedrepresentationsinCreo Parametric.This
functionmatchSimpRepItemreturns anarrayof simplifiedrepresentation
| matchingaComponentPath |     |     | for acertainfeatureas |     | wellas | the |
| ---------------------- | --- | --- | --------------------- | --- | ------ | --- |
IpfcSimpRepActionType forthatitem's actioninthe representation.Ifnone
| arefoundthe | methodprintsthe<NOTFOUND>messageand |            |     |                 |     | returnsnull. |
| ----------- | ----------------------------------- | ---------- | --- | --------------- | --- | ------------ |
| Modifying   |                                     | Simplified |     | Representations |     |              |
MethodsandPropertiesIntroduced:
• IpfcSimpRep.GetInstructions()
• IpfcSimpRep.SetInstructions()
• IpfcSimpRepInstructions.DefaultAction
• IpfcCreateNewSimpRepInstructions.NewSimpName
• IpfcSimpRepInstructions.IsTemporary
UsingtheVB API,youcan modifytheattributesof existingsimplified
representations.After youcreateorretrieve asimplifiedrepresentation,youcan
makecallsto themethodslisted in thissectionto designatenewvaluesfor the
| fieldsinthe | IpfcSimpRepInstructions |     |     | object. |     |     |
| ----------- | ----------------------- | --- | --- | ------- | --- | --- |
Tomodifyan existingsimplifiedrepresentationretrieve itandthengetthe
IpfcSimpRepInstructionsobjectbycalling
IpfcSimpRep.GetInstructions(). Ifyoucreatedthe representation
| programmaticallywithin                           |     | thesameapplication, |     | the |     |             |
| ------------------------------------------------ | --- | ------------------- | --- | --- | --- | ----------- |
| IpfcSimpRepInstructionsobjectisalreadyavailable. |     |                     |     |     |     | Onceyouhave |
modifiedthedataobject,reassignitto thecorrespondingsimplifiedrepresentation
themethodIpfcSimpRep.SetInstructions().
bycalling
| Adding     | Items | to and         | Deleting | Items | from | a   |
| ---------- | ----- | -------------- | -------- | ----- | ---- | --- |
| Simplified |       | Representation |          |       |      |     |
MethodsandPropertiesIntroduced:
• CCpfcSimpRepReverse.Create()
• CCpfcSimpRepInclude.Create()
• CCpfcSimpRepExclude.Create()
• CCpfcSimpRepSubstitute.Create()
• CCpfcSimpRepGeom.Create()
| 318 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- |

• CCpfcSimpRepGraphics.Create()
• CCpfcSimpRepDefaultEnvelope.Create()
• CCpfcSimpRepBoundBox.Create()
• CCpfcSimpRepLightWeightGraphics.Create()
• CCpfcSimpRepItem.Create()
• CCpfcSimpRep.SetInstructions()
• CCpfcSimpRepInstructions.Items()
ThemethodCCpfcSimpRepReverse.Create()
|                    |                     |             | createsan     | objectthat |
| ------------------ | ------------------- | ----------- | ------------- | ---------- |
| reversesthedefault | rulefor theassembly | componentin | thesimplified |            |
representation.
| ThemethodCCpfcSimpRepInclude.Create() |                              |     | createsan | objectthat |
| ------------------------------------- | ---------------------------- | --- | --------- | ---------- |
| includesthecomponentin                | thesimplifiedrepresentation. |     |           |            |
| ThemethodCCpfcSimpRepExclude.Create() |                              |     | createsan | objectthat |
excludesthecomponentfrom thesimplifiedrepresentation.
ThemethodCCpfcSimpRepSubstitute.Create()
|     |     |     | createsan | objectthat |
| --- | --- | --- | --------- | ---------- |
substitutesthecomponentwitha componentfrom theIpfcSubstitution
classin thesimplifiedrepresentation.
ThemethodCCpfcSimpRepGeom.Create() createsanobjectthatrepresents
thecomponentthatcontains complete geometricinformation inthe simplified
representation.
| ThemethodCCpfcSimpRepGraphics.Create() |     |     | createsan | objectthat |
| -------------------------------------- | --- | --- | --------- | ---------- |
representsthecomponentthatcontains onlydisplayinformationin thesimplified
representation.These graphicrepresentationscannotbemodifiedor referenced,
| theycan onlybe | browsedthrough. |     |     |     |
| -------------- | --------------- | --- | --- | --- |
Note
Ascomparedto graphicsrepresentations,geometryrepresentationstakelonger
| to retrieveand                                | requiremorememory. |     |     |           |
| --------------------------------------------- | ------------------ | --- | --- | --------- |
| ThemethodCCpfcSimpRepDefaultEnvelope.Create() |                    |     |     | createsan |
objectthatrepresents thecomponentin defaultenvelopetypeofsimplified
representation.
| ThemethodCCpfcSimpRepBoundBox.Create() |     |     | createsan | objectthat |
| -------------------------------------- | --- | --- | --------- | ---------- |
representsthecomponentin boundaryboxtypeofsimplifiedrepresentation.
| ThemethodCCpfcSimpRepLightWeightGraphics.Create() |     |     |     | creates |
| ------------------------------------------------- | --- | --- | --- | ------- |
anobjectthatrepresents thecomponentin lightweighttypeofsimplified
representation.Light graphicsrepresentationsof assembliescontainassembly
informationand3Dthumbnailgraphicsrepresentations ofassembly components.
SimplifiedRepresentations 319

ThemethodCCpfcSimpRepItem.Create() createsanobjectthatdefines
| thestatusofthe                          | componentor | featureofthesimplifiedrepresentation. |     |        |            |
| --------------------------------------- | ----------- | ------------------------------------- | --- | ------ | ---------- |
| ThemethodCCpfcSimpRep.SetInstructions() |             |                                       |     | setsor | changesthe |
instructionsthatspecify theinternaldatarequiredto define thesimplified
representation.
| ThemethodCCpfcSimpRepInstructions.Items |              |            |                              | appliesan | arrayof |
| --------------------------------------- | ------------ | ---------- | ---------------------------- | --------- | ------- |
| actionstothe                            | componentsor | featuresin | thesimplifiedrepresentation. |           |         |
You canaddanddeleteitems from thelist ofcomponentsina simplified
representationusingthe VBAPI. Ifyoucreatedasimplifiedrepresentationusing
theoptionExclude as thedefaultrule,youwouldgeneratealist containingthe
itemsyouwant toinclude. Similarly, ifthe defaultrulefor asimplified
representationisInclude, youcanaddthe itemsthatyouwant tobe excludedfrom
| thesimplifiedrepresentationto |           | thelist,settingthe     |     | valueofthe |     |
| ----------------------------- | --------- | ---------------------- | --- | ---------- | --- |
| EpfcSimpRepActionType         |           | toEpfcSIMPREP_EXCLUDE. |     |            |     |
| How to                        | Add Items |                        |     |            |     |
1. GettheIpfcSimpRepInstructions object,as describedin theprevious
section.
2. Specifytheactionto beapplied to theitemwith acall tooneoffollowing
methods.
3. InitializeaIpfcSimpRepItem objectfor theitem bycalling themethod
IpfcSimpRepItem.Create().
| 4. Addtheitem | tothe IpfcSimpRepItem |     |     | sequence.Putthenew |     |
| ------------- | --------------------- | --- | --- | ------------------ | --- |
IpfcSimpRepInstructionsusing
IpfcSimpRepInstructions.Items.
5. ReassigntheIpfcSimpRepInstructions objecttothe corresponding
| IpfcSimpRepobjectby |              | calling | IpfcSimpRep.SetInstructions() |     |     |
| ------------------- | ------------ | ------- | ----------------------------- | --- | --- |
| How to              | Remove Items |         |                               |     |     |
Followtheprocedure above,exceptremovetheunwanted IpfcSimpRepItem
fromthesequence.
| Simplified | Representation |     |     | Utilities |     |
| ---------- | -------------- | --- | --- | --------- | --- |
MethodsIntroduced:
• IpfcModelItemOwner.ListItems()
• IpfcModelItemOwner.GetItemById()
• IpfcSolid.GetSimpRep()
| 320 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

• IpfcSolid.SelectSimpRep()
• IpfcSolid.ActivateSimpRep()
• IpfcSolid.GetActiveSimpRep()
Thissectiondescribesthe utilitymethods thatrelatetosimplifiedrepresentations.
| ThemethodIpfcModelItemOwner.ListItems()   |            |     | can listall  | ofthe |
| ----------------------------------------- | ---------- | --- | ------------ | ----- |
| simplifiedrepresentations                 | ina Solid. |     |              |       |
| ThemethodIpfcModelItemOwner.GetItemById() |            |     | initializesa |       |
IpfcSimpRepobject.Ittakesan integerid.
Note
theVB APIsupports simplifiedrepresentationofAssembliesonly,not Parts.
ThemethodIpfcSolid.GetSimpRep() initializesaIpfcSimpRep object.
ThemethodtakesSimpRepnameas theinputargument. SimpRepnameisthe name
ofthesimplifiedrepresentationin thesolid. Ifyouspecify thisargument, the
| methodignoresthe rep_id. |     |     |     |     |
| ------------------------ | --- | --- | --- | --- |
ThemethodIpfcSolid.SelectSimpRep() createsaCreo Parametricmenu
toenableinteractiveselection.Themethod takestheowningsolidasinput,and
outputsthe objecttothe selectedsimplifiedrepresentation.Ifyouchoose theQuit
| menubutton,themethod                   | throwsan exception | XToolkitUserAbort |     |     |
| -------------------------------------- | ------------------ | ----------------- | --- | --- |
| ThemethodsIpfcSolid.GetActiveSimpRep() |                    |                   | and |     |
IpfcSolid.ActivateSimpRep()enableyouto findandget thecurrently
activesimplifiedrepresentation,respectively.Givenan assembly object,
IpfcSolid.GetActiveSimpRep()returns
|     |     | theobjectto | thecurrentlyactive |     |
| --- | --- | ----------- | ------------------ | --- |
simplifiedrepresentation.Ifthe currentrepresentationisthemasterrepresentation,
thereturnisnull.
| ThemethodIpfcSolid.ActivateSimpRep() |     | activatestherequested |     |     |
| ------------------------------------ | --- | --------------------- | --- | --- |
simplifiedrepresentation.
Toseta simplifiedrepresentation tobe thecurrentlydisplayedmodel, youmust
alsocallIpfcModel.Display().
SimplifiedRepresentations 321

25
| Task | Based | Application | Libraries |
| ---- | ----- | ----------- | --------- |
ManagingApplicationArguments.............................................................................323
LaunchingaCreoTOOLKITDLL..............................................................................324
LaunchingTasksfromJ-LinkTaskLibraries...............................................................326
| Applicationscreatedusing | differentCreo | ParametricAPI | productsare |
| ------------------------ | ------------- | ------------- | ----------- |
interoperable.These productsuseCreo Parametricasthe mediumof interaction,
eliminatingthetaskof writingnative-platformspecificinteractions between
| differentprogramming | languages. |     |     |
| -------------------- | ---------- | --- | --- |
ApplicationinteroperabilityallowstheVB APIapplicationsto callintoCreo
TOOLKIT fromareasnotcoveredin thenativeinterface.It allowsyouto puta
VBAor VB.NETfront endon legacyCreo TOOLKITapplications,andalso
allowsyouto useapplicationsandlisteners inconjunction witha orasynchronous
application.
| 322 |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | ------------------------------- | --- |

Managing Application Arguments
TheVB APIpasses applicationdatato andfrom tasksin otherapplicationsas
membersof asequenceof IpfcArgumentobjects.Applicationarguments
consistofa labelanda value.Thevalue maybeof anyoneofthe followingtypes:
• Integer
• Double
• Boolean
• ASCIIstring(a non-encodedstring,provided forcompatibility witharguments
providedfrom Capplications)
• String(a fullyencodedstring)
• IpfcSelection (aselectionof anitem in aCreo Parametricsession)
• IpfcTransform3D (acoordinatesystemtransformation matrix)
MethodsandPropertiesIntroduced:
• CMpfcArgument.CreateIntArgValue()
• CMpfcArgumentCreateDoubleArgValue()
• CMpfcArgumentCreateBoolArgValue()
• CMpfcArgumentCreateASCIIStringArgValue()
• CMpfcArgumentCreateStringArgValue()
• CMpfcArgumentCreateSelectionArgValue()
• CMpfcArgumentCreateTransformArgValue()
• IpfcArgValue.discr
• IpfcArgValue.IntValue
• IpfcArgValue.DoubleValue
• IpfcArgValue.BoolValue
• IpfcArgValue.ASCIIStringValue
• IpfcArgValue.StringValue
• IpfcArgValue.SelectionValue
• IpfcArgValue.TransformValue
Theclass IpfcArgValuecontainsone oftheseventypesofvalues. theVB API
providesdifferent methodsto createeachofthe seventypesof argumentvalues.
ThepropertyIpfcArgValue.discr returnsthetypeofvalue containedin the
argumentvalue object.
Usethe methodslistedabove toaccess andmodifytheargumentvalues.
TaskBasedApplicationLibraries 323

| Modifying | Arguments |     |     |     |
| --------- | --------- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcArgument.Create()
• IpfcpfcArgument.Label
• IpfcpfcArgument.Value
ThemethodCCpfcArgument.Create() createsanewargument. Provide a
| nameandvalueas | theinputargumentsof |     | thismethod. |     |
| -------------- | ------------------- | --- | ----------- | --- |
ThepropertyIpfcpfcArgument.Label
|     |     |     | returns | thelabelof theargument. |
| --- | --- | --- | ------- | ----------------------- |
ThepropertyIpfcpfcArgument.Value returns thevalueof theargument.
| Launching | a Creo | TOOLKIT |     | DLL |
| --------- | ------ | ------- | --- | --- |
Themethodsdescribed in thissectionenablea theVB APIuserto registerand
launcha CreoTOOLKITDLL from anapplication. Theability tolaunchand
| controlaCreoTOOLKIT | applicationenablesthefollowing: |     |       |                  |
| ------------------- | ------------------------------- | --- | ----- | ---------------- |
| • Reuseofexisting   | CreoTOOLKITcodewith             |     | theVB | APIapplications. |
• ATBoperations.
MethodsandPropertiesIntroduced:
• IpfcBaseSession.LoadProToolkitDll()
• IpfcBaseSession.LoadProToolkitLegacyDll()
• IpfcBaseSession.GetProToolkitDll()
• IpfcDll.ExecuteFunction()
• IpfcDll.Id
• IpfcDll.IsActive()
• IpfcDll.Unload()
Usethe methodIpfcBaseSession.LoadProToolkitDll() toregister
andstarta CreoTOOLKITDLL. Theinput parametersof thismethodaresimilar
| tothe fieldsofaregistryfileandareas |        |                    | follows:    |             |
| ----------------------------------- | ------ | ------------------ | ----------- | ----------- |
| • ApplicationName—Thename           |        | oftheapplicationto |             | initialize. |
| • DllPath—Thefull                   | pathto | theDLL             | binaryfile. |             |
• TextPath—Thepathto theapplication’s messageanduserinterfacetextfiles.
• UserDisplay—Setthisparameter to truetoregisterthe applicationin the
CreoParametricuserinterface andto seeerror messagesif theapplication
fails.Ifthis parameteris false,theapplicationwillbeinvisibleto the user.
| 324 |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

Theapplication'suser_initialize() functionis calledwhenthe application
isstarted.The methodreturns ahandletothe loadedCreoTOOLKITDLL.
Inorder toregisterandstarta legacyPro/TOOLKITDLL thatisnotUnicode-
compliant,usethemethod
| IpfcBaseSession.LoadProToolkitLegacyDll().This |     |     |     | method |
| ---------------------------------------------- | --- | --- | --- | ------ |
conveysto Creo ParametricthattheloadedDLLapplicationisnot Unicode-
compliantandbuiltin thepre-Wildfire4.0environment. Ittakesthe sameinput
| parametersas | theearliermethod |     |     |     |
| ------------ | ---------------- | --- | --- | --- |
IpfcBaseSession.LoadProToolkitDll().
Note
The methodIpfcBaseSession.LoadProToolkitLegacyDll()
mustbe usedonlyby apre-Wildfire4.0 applicationto loadapre-Wildfire4.0
| Pro/TOOLKIT |                                          | DLL. |     |            |
| ----------- | ---------------------------------------- | ---- | --- | ---------- |
| Usethe      | methodIpfcBaseSession.GetProToolkitDll() |      |     | to obtaina |
CreoTOOLKITDLL handle.Specifythe Application_Id,thatis,theDLL’s
identifierstringas theinputparameter ofthismethod.The methodreturns the
DLLobjectornullif theDLL wasnot insession. The Application_Idcan be
determinedasfollows:
• Usethe functionProToolkitDllIdGet() within theDLL applicationto
NULLto
| getastringrepresentation |     | oftheDLL | application.Pass | thefirst |
| ------------------------ | --- | -------- | ---------------- | -------- |
argument ofProToolkitDllIdGet() toget thestringidentifier forthe
callingapplication.
• Usethe Getmethodfor theId attribute inthe DLLinterface.The method
| IpfcDll.Idreturnsthe |     | DLLidentifier | string. |     |
| -------------------- | --- | ------------- | ------- | --- |
Usethe methodIpfcDll.ExecuteFunction()to callaproperlydesignated
functionin theCreo TOOLKIT DLLlibrary. Theinputparametersof thismethod
are:
| • FunctionName—Nameof |     | thefunctionin | theCreo TOOLKITDLL |     |
| --------------------- | --- | ------------- | ------------------ | --- |
application.
• InputArguments—Input arguments tobe passedto thelibraryfunction.
Themethodreturns anobjectof IpfcFunctionReturn.This interface
containsdatareturnedby aCreo TOOLKIT functioncall.The objectcontainsthe
returnvalue,asinteger, oftheexecutedfunction andtheoutput argumentspassed
| backfrom                    | thefunction | call.     |                        |         |
| --------------------------- | ----------- | --------- | ---------------------- | ------- |
| ThemethodIpfcDll.IsActive() |             |           | determineswhetheraCreo | TOOLKIT |
| DLLpreviouslyloadedby       |             | themethod |                        |         |
IpfcBaseSession.LoadProToolkitDll()is stillactive.
TaskBasedApplicationLibraries 325

| ThemethodIpfcDll.Unload()                 |     |           | isusedto | shutdownaCreo |      | TOOLKITDLL         |
| ----------------------------------------- | --- | --------- | -------- | ------------- | ---- | ------------------ |
| previouslyloadedby                        |     | themethod |          |               |      |                    |
| IpfcBaseSession.LoadProToolkitDll()andthe |     |           |          |               |      | application'suser_ |
| terminate()function                       |     | iscalled. |          |               |      |                    |
| Launching                                 |     | Tasks     | from     | J-Link        | Task |                    |
Libraries
Themethodsdescribed in thissectionallowyouto launchtasks fromapredefined
J-Linktasklibrary.
MethodsIntroduced:
• IpfcBaseSession.StartJLinkApplication()
• IpfcJLinkApplication.ExecuteTask()
• IpfcJLinkApplication.IsActive()
• IpfcJLinkApplication.Stop()
methodIpfcBaseSession.StartJLinkApplication()
Usethe tostart
aapplication.The inputparametersofthismethod aresimilarto thefieldsof a
| registryfileandareas |     | follows: |     |     |     |     |
| -------------------- | --- | -------- | --- | --- | --- | --- |
• ApplicationName—Assignsa uniquename tothis J-Linkapplication.
| • ClassName—Specifiesthename |     |     | ofthe | Javaclass | thatcontainsthe |     |
| ---------------------------- | --- | --- | ----- | --------- | --------------- | --- |
application’s startand stopmethod.Thisshould bea fullyqualifiedJava
| packageandclass |     | name. |     |     |     |     |
| --------------- | --- | ----- | --- | --- | --- | --- |
• StartMethod—Specifiesthestartmethodofthe J-Linkapplication.
| • StopMethod—Specifiesthe |     |     | stopmethodof | theJ-Linkapplication. |     |     |
| ------------------------- | --- | --- | ------------ | --------------------- | --- | --- |
• AdditionalClassPath—Specifiesthelocationsof packagesandclasses that
mustbeloadedwhen startingthisJ-Link application.Ifthisparameter is
| specifiedasnull,the |     | defaultclasspath |     | locationsareused. |     |     |
| ------------------- | --- | ---------------- | --- | ----------------- | --- | --- |
• TextPath—Specifiestheapplicationtextpathformenus andmessages.Ifthis
| parameteris | specifiedasnull,the |     | defaulttextlocationsareused. |     |     |     |
| ----------- | ------------------- | --- | ---------------------------- | --- | --- | --- |
• UserDisplay—Specifieswhetherto displaytheapplicationin theAuxiliary
| Applicationsdialogboxin |     |     | CreoParametric. |     |     |     |
| ----------------------- | --- | --- | --------------- | --- | --- | --- |
Uponstartingtheapplication, thestaticstart()method isinvoked.Themethod
returnsaIpfcJLinkApplication referringto theJ-Linkapplication.
| ThemethodIpfcJLinkApplication.ExecuteTask() |     |     |     |     |     | callsaregistered |
| ------------------------------------------- | --- | --- | --- | --- | --- | ---------------- |
taskmethodin theJ-Link application.Theinput parametersof thismethodare:
| • Nameofthe | taskto | be executed. |     |     |                                 |     |
| ----------- | ------ | ------------ | --- | --- | ------------------------------- | --- |
| 326         |        |              |     |     | Creo®ParametricVBAPIUser’sGuide |     |

• Asequenceof namevaluepair argumentscontainedby theinterface
IpfcArguments.
| Themethodoutputs                   | an arrayofoutput | arguments. |                   |       |
| ---------------------------------- | ---------------- | ---------- | ----------------- | ----- |
| ThemethodJLinkApplication.IsActive |                  |            | returnsaTruevalue | ifthe |
applicationspecifiedby theJLink.IpfcJLinkApplication objectis
active.
| ThemethodIpfcJLinkApplication.Stop() |     |     | stopstheapplication |     |
| ------------------------------------ | --- | --- | ------------------- | --- |
specifiedby theIpfcJLinkApplication object.Thismethodactivates the
application’sstaticStop()method.
| TaskBasedApplicationLibraries |     |     |     | 327 |
| ----------------------------- | --- | --- | --- | --- |

26
Graphics
Overview................................................................................................................329
GettingMouseInput................................................................................................329
DisplayingGraphics.................................................................................................330
DisplayListsandGraphics.......................................................................................333
Thischaptercoversthe VBAPI Graphicsincludingdisplayinglists, displaying
textand usingthe mouse.
328 Creo®ParametricVBAPIUser’sGuide

Overview
Themethodsdescribed in thissectionallowyouto draw temporarygraphicsin a
displaywindow.Methods thatareidentifiedas2D areusedto drawentities(arcs,
polygons,andtext) inscreen coordinates.Otherentitiesmaybe drawn using the
currentmodel’scoordinatesystem orthescreen coordinatesystem’s lines, circles,
andpolylines.Methods arealso includedfor manipulating textpropertiesand
accessingmouseinputs.
Getting Mouse Input
Thefollowingmethodsareused toread themouseposition in screencoordinates
withthemouse buttondepressed.Eachmethod outputsthe positionandan
enumeratedtypedescriptionof whichmousebutton waspressedwhen themouse
wasatthatposition.Thesevalues arecontainedin theinterface
IpfcMouseStatus.
Theenumeratedvalues aredefinedin EpfcMouseButton andareasfollows:
• EpfcMOUSE_BTN_LEFT
• EpfcMOUSE_BTN_RIGHT
• EpfcMOUSE_BTN_MIDDLE
• EpfcMOUSE_BTN_LEFT_DOUBLECLICK
MethodsIntroduced:
• IpfcSession.UIGetNextMousePick()
• IpfcSession.UIGetCurrentMouseStatus()
ThemethodIpfcSession.UIGetNextMousePick() returns themouse
positionwhenyoupressa mousebutton.The inputargumentisthe mousebutton
thatyouexpectthe userto select.
ThemethodIpfcSession.UIGetCurrentMouseStatus() returnsa
valuewheneverthe mouseismoved ora buttonis pressed.Withthis methoda
buttondoesnot haveto be pressedfor avalueto be returned.You canusean input
argumentto flag whetherornotthereturned positionsaresnapped to thewindow
grid.
Drawing a Mouse Box
Thismethodallows youtodraw amousebox.
MethodIntroduced:
• IpfcSession.UIPickMouseBox()
Graphics 329

ThemethodIpfcSession.UIPickMouseBox() draws adynamicrectangle
fromaspecifiedpoint inscreencoordinatesto thecurrentmouseposition untilthe
userpressesthe leftmouse button.The returnvaluefor thismethodis ofthe type
IpfcOutline3D.
You cansupplythefirst cornerlocationprogrammaticallyoryoucan allowthe
| usertoselectboth | cornersofthe | box. |     |     |
| ---------------- | ------------ | ---- | --- | --- |
| Displaying       | Graphics     |      |     |     |
Allthemethodsin thissection drawgraphicsin theCreo Parametriccurrent
| windowandusethecolorandlinestylesetby |     | callsto |     |     |
| ------------------------------------- | --- | ------- | --- | --- |
IpfcBaseSession.SetStdColorFromRGB()
and
| IpfcBaseSession.SetLineStyle().The |     |     | methodsdrawthe | graphicsin |
| ---------------------------------- | --- | --- | -------------- | ---------- |
theCreoParametricgraphics color.The defaultgraphicscoloris white.
Themethodsin this sectionarecalledusingthe interfaceIpfcDisplay.This
interfaceisextendedbythe IpfcBaseSession interface.Thisarchitecture
| allowsyouto | callallthesemethodson | anyIpfcSession | object. |     |
| ----------- | --------------------- | -------------- | ------- | --- |
MethodsIntroduced:
• IpfcDisplay.SetPenPosition()
• IpfcDisplay.DrawLine()
• IpfcDisplay.DrawPolyline()
• IpfcDisplay.DrawCircle()
• IpfcDisplay.DrawArc2D()
• IpfcDisplay.DrawPolygon2D()
ThemethodIpfcDisplay.SetPenPosition() setsthe pointatwhichyou
functionIpfcDisplay.DrawLine()
| wantto startdrawinga | line.The |     |     | draws |
| -------------------- | -------- | --- | --- | ----- |
alinetothe givenpointfrom theposition givenin thelastcallto eitherofthe two
| functions.CallIpfcDisplay.SetPenPosition() |     |     | forthe | startofthe |
| ------------------------------------------ | --- | --- | ------ | ---------- |
polyline,andIpfcDisplay.DrawLine() for eachvertex.Ifyouusethese
methodsin two-dimensionalmodes,usescreen coordinatesinsteadof solid
coordinates.
ThemethodIpfcDisplay.DrawCircle()
usessolidcoordinatesforthe
centerof thecircleandtheradiusvalue. Thecirclewillbe placedtothe XYplane
ofthemodel.
ThemethodIpfcDisplay.DrawPolyline() also drawspolylines, usingan
| arrayto define | thepolyline. |     |     |     |
| -------------- | ------------ | --- | --- | --- |
Intwo-dimensionalmodelstheDisplay Graphicsmethodsdrawgraphics atthe
| specifiedscreen | coordinates. |     |                                 |     |
| --------------- | ------------ | --- | ------------------------------- | --- |
| 330             |              |     | Creo®ParametricVBAPIUser’sGuide |     |

| ThemethodIpfcDisplay.DrawPolygon2D() |     |     |     |     | drawsapolygonin |     | screen |
| ------------------------------------ | --- | --- | --- | --- | --------------- | --- | ------ |
coordinates.The methodIpfcDisplay.DrawArc2D()draws anarcin screen
coordinates.
| Controlling | Graphics |     | Display |     |     |     |     |
| ----------- | -------- | --- | ------- | --- | --- | --- | --- |
PropertiesIntroduced:
• IpfcDisplay.CurrentGraphicsColor
• IpfcDisplay.CurrentGraphicsMode
| ThepropertyIpfcDisplay.CurrentGraphicsColor |     |     |     |     |     | returns | theCreo |
| ------------------------------------------- | --- | --- | --- | --- | --- | ------- | ------- |
Parametricstandardcolorusedto displaygraphics.The Creo Parametricdefaultis
EpfcCOLOR_DRAWING(white).
ThepropertyIpfcDisplay.CurrentGraphicsMode returnsthemode used
todraw graphics:
• EpfcDRAW_GRAPHICS_NORMAL—Creo Parametricdrawsgraphics inthe
| requiredcolorin                     | eachinvocation. |     |     |     |                         |     |     |
| ----------------------------------- | --------------- | --- | --- | --- | ----------------------- | --- | --- |
| • EpfcDRAW_GRAPHICS_COMPLEMENT—Creo |                 |     |     |     | Parametricdrawsgraphics |     |     |
normally,but willerasegraphicsdrawn asecondtime in thesamelocation.
| Thisallows | youtocreaterubberband |          |     | lines.   |     |        |     |
| ---------- | --------------------- | -------- | --- | -------- | --- | ------ | --- |
| Example    | Code:                 | Creating |     | Graphics | On  | Screen |     |
Thesamplecodein thefilepfcGraphicsExamples.vblocated at<creo_
| vbapi_loadpoint>/vbapi_examplesdemonstratesthe |              |     |                 |     |              | useofmouse- |     |
| ---------------------------------------------- | ------------ | --- | --------------- | --- | ------------ | ----------- | --- |
| trackingmethodsto                              | drawgraphics |     | onthescreen.The |     | staticmethod |             |     |
DrawRubberbandLinepromptstheuserto pickascreenpoint.The example
usesthe ‘complementmode’to causethelineto displayanderaseasthe user
| movesthemousearound |     | thewindow. |     |     |     |     |     |
| ------------------- | --- | ---------- | --- | --- | --- | --- | --- |
Note
| This exampleusesthe                    |     | methodtransformPosition |     |     |              | to convertthe |     |
| -------------------------------------- | --- | ----------------------- | --- | --- | ------------ | ------------- | --- |
| coordinatesintothe3Dcoordinatesystemof |     |                         |     |     | asolidmodel, | ifoneis       |     |
displayed.
| Displaying | Text | in the | Graphics |     | Window |     |     |
| ---------- | ---- | ------ | -------- | --- | ------ | --- | --- |
MethodIntroduced:
• IpfcDisplay.DrawText2D()
| Graphics |     |     |     |     |     |     | 331 |
| -------- | --- | --- | --- | --- | --- | --- | --- |

ThemethodIpfcDisplay.DrawText2D() places textatapositionspecified
inscreencoordinates. Ifyouwant toadd texttoa particularposition onthe solid,
youmusttransform thesolidcoordinatesintoscreencoordinatesbyusing the
viewmatrix.
Textitems drawn arenotknown to CreoParametricandtherefore arenotredrawn
whenyouselectView,Repaint.Tonotify theCreoParametricoftheseobjects,
createthem insidethe OnDisplay()method oftheDisplay Listener.
| Controlling | Text Attributes |     |     |     |
| ----------- | --------------- | --- | --- | --- |
PropertiesIntroduced:
• IpfcDisplay.extHeight
• IpfcDisplay.WidthFactor
• IpfcDisplay.RotationAngle
• IpfcDisplay.SlantAngle
| Thesepropertiescontroltheattributesof |     |     | textaddedby | callsto |
| ------------------------------------- | --- | --- | ----------- | ------- |
IpfcDisplay.DrawText2D().
| You canaccessthe | followinginformation: |     |     |     |
| ---------------- | --------------------- | --- | --- | --- |
• Text height(inscreencoordinates)
• Widthratioof eachcharacter, includingthegap,as aproportionofthe height
| • Rotationangleof | thewholetext, |                     | incounterclockwise | degrees |
| ----------------- | ------------- | ------------------- | ------------------ | ------- |
| • Slantangleof    | thetext,      | in clockwisedegrees |                    |         |
| Controlling       | Text          | Fonts               |                    |         |
MethodsandPropertiesIntroduced:
• IpfcDisplay.DefaultFont
• IpfcDisplay.CurrentFont
• IpfcDisplay.GetFontById()
• IpfcDisplay.GetFontByName()
| ThepropertyIpfcDisplay.DefaultFont |     |     | returns | thedefaultCreo |
| ---------------------------------- | --- | --- | ------- | -------------- |
Parametrictextfont. Thetextfonts areidentifiedin Creo Parametricbynames and
| byintegeridentifiers. | Tofind | aspecific | font,usethemethods |     |
| --------------------- | ------ | --------- | ------------------ | --- |
IpfcDisplay.GetFontById()orIpfcDisplay.GetFontByName().
| 332 |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

Display Lists and Graphics
Whengeneratinga displayofasolidin awindow, Creo Parametricmaintainstwo
displaylists. Adisplaylistcontains asetofvectors thatareusedto representthe
shapeof thesolidin theview.A3D displaylist containsa setofthree-
dimensionalvectors thatrepresentan approximationto the geometryoftheedges
ofthesolid. Thislist getsrebuilt everytimethesolidis regenerated.
A2Ddisplaylistcontains thetwo-dimensionalprojectionsofthe edgesofthe
solid3D displaylistonto thecurrentwindow. Itis rebuiltfromthe 3Ddisplaylist
whentheorientation ofthesolidchanges. Themethods inthis sectionenableyou
toaddyour ownvectorsto thedisplaylists, sothatthegraphicswillbe
redisplayedautomatically byCreo Parametricuntilthe displaylists arerebuilt.
Whenyouaddgraphicsitems tothe 2Ddisplaylist,theywillbe regeneratedafter
eachrepaint (whenzooming andpanning) andwillbe includedin plotscreatedby
CreoParametric.Whenyouaddgraphicsto the 3Ddisplaylist,yougetthe further
benefitthatthegraphicssurvive achangetothe orientationofthe solidandare
displayedevenwhen youspinthesoliddynamically.
MethodsIntroduced:
• IpfcDisplayListener.OnDisplay()
• IpfcDisplay.CreateDisplayList2D()
• IpfcDisplay.CreateDisplayList3D()
• IpfcDisplayList2D.Display()
• IpfcDisplayList3D.Display()
• IpfcDisplayList2D.Delete()
• IpfcDisplayList3D.Delete()
Adisplaylistener isaclass thatacts similarlyto anactionlistener. You must
implementthemethod inheritedfrom theIpfcDisplayListener interface.
Theimplementationshouldprovidecallsto methods ontheprovided
IpfcDisplayobjectto produce2Dor 3Dgraphics.
Inorder tocreateadisplaylist in CreoParametric,youcall
IpfcDisplay.CreateDisplayList2Dor
IpfcDisplay.CreateDisplayList3D()to tellCreo Parametrictouse
yourlistenerto createthedisplaylist vectors.
IpfcDisplayList2D.Display()or
IpfcDisplayList3D.Display()willdisplayorredisplay theelementsin
yourdisplaylist.Theapplicationshoulddeletethe displaylistdatawhen itis no
longerneeded.
ThemethodIpfcDisplayList2D.Delete() andthe method
IpfcDisplayList3D.Delete()willremoveboth thespecifieddisplaylist
fromasession.
Graphics 333

Note
The methodIpfcWindow.Refresh()does notcauseeitherofthedisplay
lists tobe regenerated,butsimplyrepaintsthe windowusingthe 2Ddisplay
list.
Exceptions
Possibleexceptions thatmightbethrownby displayinggraphicsmethods are
shownin thefollowingtable:
| Exception            |     | Reason                 |     |
| -------------------- | --- | ---------------------- | --- |
| IpfcXToolkitNotExist |     | Thedisplaylistisempty. |     |
IpfcXToolkitNotFound
Thefunctioncouldnotfindthedisplaylistorthe
fontspecifiedinapreviouscallto
IpfcDisplay.CurrentFontwasnotfound.
IpfcXToolkitCantOpen
Theuseofdisplaylistsisdisabled.
| IpfcXToolkitAbort    |     | Thedisplaywasaborted.             |     |
| -------------------- | --- | --------------------------------- | --- |
| IpfcXToolkitNotValid |     | Thespecifieddisplaylistisinvalid. |     |
IpfcXToolkitInvalidItem Thereisaninvaliditeminthedisplaylist.
IpfcXToolkitGeneralError Thespecifieddisplaylistisalreadyintheprocessof
beingdisplayed.
Example Code
Thesamplecodein thefilepfcGraphicsExamples.vblocated at<creo_
| vbapi_loadpoint>/vbapi_examplesdemonstratesthe |                      |                  | useof       |
| ---------------------------------------------- | -------------------- | ---------------- | ----------- |
| IpfcDisplaymethods                             | with3D displaylists. | Thestaticmethod  |             |
| AddCircleDisplay()createsanew3D                |                      | displaylistwhose | graphicsare |
generatedbythe codein the OnDisplay()methodof theDisplay Circlesclass.
Thisdisplaylist placescirclesatall ofthevertices ofapart modelonthe screen.
| 334 |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | ------------------------------- | --- |

27
External Data
ExternalData..........................................................................................................336
Exceptions..............................................................................................................339
Thischapterexplains usingExternalDatain theVB API.
335

External Data
Thischapterdescribes howtostore andretrieve externaldata. Externaldata
enablesatheVB APIapplicationto storeitsowndatain aCreo Parametric
databasein suchawaythatitis invisibletothe CreoParametricuser.This method
isdifferentfrom othermeansof storageaccessible through theCreo Parametric
userinterface.
Introduction to External Data
Externaldataprovidesa wayforthe CreoParametricapplicationto storeitsown
privateinformationabouta CreoParametricmodelwithin themodelfile.The data
isbuiltandinterrogated by theapplicationas aworkspace datastructure. Itis
savedto themodelfilewhenthe modelissaved, andretrievedwhenthe modelis
retrieved.The externaldataisotherwiseignoredby Creo Parametric;the
applicationhascomplete controloverform andcontent.
Theexternaldatafora specificCreo Parametricmodelisbrokendown into
classesandslots.Aclass isanamed‘‘bin’’foryourdata, andidentifies itasyours
sonootherCreo ParametricAPIapplication(or otherclassesin yourown
application)willuseitbymistake. Anapplicationusuallyneedsonlyoneclass.
Theclass nameshouldbe unique foreachapplicationanddescribethe roleofthe
datain yourapplication.
Eachclasscontains asetofdataslots. Eachslotisidentifiedby an identifierand
optionally, aname.Aslot containsasingle dataitemof oneofthe following
types:
TheVBAPIType Data
EpfcEXTDATA_INTEGER integer
EpfcEXTDATA_DOUBLE double
EpfcEXTDATA_STRING string
ThetheVB API interfacesusedto accessexternaldatainCreo Parametricare:
TheVBAPIType DataType
IpfcExternalDataAccess Thisisthetoplevelobjectandis
createdwhenattemptingtoaccess
externaldata.
IpfcExternalDataClass Thisisaclassofexternaldataandis
identifiedbyauniquename.
IpfcExternalDataSlot Thisisacontainerforoneitemof
data.Eachslotisstoredinaclass.
IpfcExternalData Thisisacompactdatastructurethat
containseitheraninteger,doubleor
stringvalue.
336 Creo®ParametricVBAPIUser’sGuide

| Compatibility |     |     | with | Creo | TOOLKIT |     |     |     |     |     |
| ------------- | --- | --- | ---- | ---- | ------- | --- | --- | --- | --- | --- |
theVBAPI andCreo TOOLKIT shareexternaldatain thesamemanner.theVB
APIexternaldataisaccessible byCreo TOOLKIT andthereverse isalso true.
However,an errorwillresultif theVB APIattempts toaccess externaldata
| previouslystored |     | byCreo   | TOOLKIT |      | asastream. |     |     |     |     |     |
| ---------------- | --- | -------- | ------- | ---- | ---------- | --- | --- | --- | --- | --- |
| Accessing        |     | External |         | Data |            |     |     |     |     |     |
MethodsIntroduced:
• IpfcModel.AccessExternalData()
• IpfcModel.TerminateExternalData()
• IpfcExternalDataAccess.IsValid()
| ThemethodIpfcModel.AccessExternalData() |     |                     |     |     |     |            |           | preparesCreo |     |     |
| --------------------------------------- | --- | ------------------- | --- | --- | --- | ---------- | --------- | ------------ | --- | --- |
| Parametrictoread                        |     | externaldatafromthe |     |     |     | modelfile. | Itreturns | the          |     |     |
IpfcExternalDataAccessobjectthatisused toread andwritedata. This
| methodshouldbe                             |     | calledonlyoncefor |     |     | anygivenmodelin |     |     | session.  |     |     |
| ------------------------------------------ | --- | ----------------- | --- | --- | --------------- | --- | --- | --------- | --- | --- |
| ThemethodIpfcModel.TerminateExternalData() |     |                   |     |     |                 |     |     | stopsCreo |     |     |
Parametricfromaccessing externaldataina model.Whenyouusethis methodall
externaldatain themodelwillbe removed.Permanentremoval willoccurwhen
themodelissaved.
Note
Ifyouneedto preservetheexternaldatacreatedin session,you mustsave the
| modelbeforecalling                        |          |     | thisfunction.Otherwise,your |     |     |        |        | datawillbe    | lost. |     |
| ----------------------------------------- | -------- | --- | --------------------------- | --- | --- | ------ | ------ | ------------- | ----- | --- |
| ThemethodIpfcExternalDataAccess.IsValid() |          |     |                             |     |     |        |        | determinesif  |       | the |
| IpfcExternalDataAccessobjectcan           |          |     |                             |     |     | beused | toread | andwritedata. |       |     |
| Storing                                   | External |     | Data                        |     |     |        |        |               |       |     |
MethodsandPropertiesIntroduced:
• IpfcExternalDataAccess.CreateClass()
• IpfcExternalDataClass.CreateSlot()
• IpfcExternalDataSlot.Value
Thefirststep in storingexternaldatain anewclass andslotis tosetup aclass
| usingthe | methodIpfcExternalDataAccess.CreateClass(), |     |     |     |     |     |     |     |     | which |
| -------- | ------------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- | ----- |
providesthe classname. Themethodoutputs IpfcExternalDataClass,
| usedbythe    | applicationto |     | referencetheclass. |     |     |     |     |     |     |     |
| ------------ | ------------- | --- | ------------------ | --- | --- | --- | --- | --- | --- | --- |
| ExternalData |               |     |                    |     |     |     |     |     |     | 337 |

| Thenextstepis                  | touseIpfcExternalDataClass.CreateSlot() |            |             |                  |          |     | to  |
| ------------------------------ | --------------------------------------- | ---------- | ----------- | ---------------- | -------- | --- | --- |
| createan emptydataslotandinput |                                         | aslotname. |             | Themethodoutputs |          | a   |     |
| IpfcExternalDataSlotobjectto   |                                         |            | identifythe |                  | newslot. |     |     |
Note
| Slot namescannotbeginwith |     | anumber. |     |     |     |     |     |
| ------------------------- | --- | -------- | --- | --- | --- | --- | --- |
ThepropertyIpfcExternalDataSlot.Value specifies thedatatypeofa
| slotandwritesan                                   | itemof thattypetothe              |     | slot.Theinputis |     | a             |            |     |
| ------------------------------------------------- | --------------------------------- | --- | --------------- | --- | ------------- | ---------- | --- |
| IpfcExternalDataobjectthatyoucancreateby          |                                   |     |                 |     | callinganyone | ofthe      |     |
| methodsin thenextsection.                         |                                   |     |                 |     |               |            |     |
| Example                                           | code                              |     |                 |     |               |            |     |
| Thesamplecodein                                   | thefilepfcExternalDataExamples.vb |     |                 |     |               | locatedat  |     |
| <creo_vbapi_loadpoint>/vbapi_examplesdemonstrates |                                   |     |                 |     |               | theusageof |     |
externaldata. Itprovidesutility methodsto convertaVB hashtableto amodel's
externaldata.
| Initializing | Data Objects |     |     |     |     |     |     |
| ------------ | ------------ | --- | --- | --- | --- | --- | --- |
MethodsIntroduced:
• CMpfcExternal.CreateIntExternalData()
• CMpfcExternal.CreateDoubleExternalData()
• CMpfcExternal.CreateStringExternalData()
Thesemethodsinitialize aIpfcExternalDataobjectwith theappropriate
datainputs.
| Retrieving | External | Data |     |     |     |     |     |
| ---------- | -------- | ---- | --- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• IpfcExternalDataAccess.LoadAll()
• IpfcExternalDataAccess.ListClasses()
• IpfcExternalDataClass.ListSlots()
• IpfcExternalData.discr
• IpfcExternalData.IntegerValue
• IpfcExternalData.DoubleValue
• IpfcExternalData.StringValue
| 338 |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | --- | ------------------------------- | --- | --- |

Forimprovedperformance,externaldataisnotloadedautomaticallyintomemory
| withthemodel. | Whenthe | modelis insession, | callthe | method |     |
| ------------- | ------- | ------------------ | ------- | ------ | --- |
IpfcExternalDataAccess.LoadAll() toretrieve allthe externaldatafor
thespecifiedmodelfrom theCreo Parametricmodelfileandputitin the
workspace.Themethodneedsto be calledonly onceto retrieve allthedata.
| ThemethodIpfcExternalDataAccess.ListClasses()   |     |                       |                 |     | returnsa   |
| ----------------------------------------------- | --- | --------------------- | --------------- | --- | ---------- |
| sequenceofIpfcExternalDataClasses               |     |                       | registeredinthe |     | model.The  |
| methodIpfcExternalDataClass.ListSlots()providea |     |                       |                 |     | sequenceof |
| IpfcExternalDataSlots                           |     | existingfor           | eachclass.      |     |            |
| Tofind outadatatypeofa                          |     | IpfcExternalData,call |                 |     |            |
IpfcExternalData.discrandthencalloneof thesepropertiestoget the
| data,dependingon | thedatatype: |     |     |     |     |
| ---------------- | ------------ | --- | --- | --- | --- |
• IpfcExternalData.IntegerValue
• IpfcExternalData.DoubleValue
• IpfcExternalData.StringValue
| Example                                           | Code                              |     |     |     |            |
| ------------------------------------------------- | --------------------------------- | --- | --- | --- | ---------- |
| Thesamplecodein                                   | thefilepfcExternalDataExamples.vb |     |     |     | locatedat  |
| <creo_vbapi_loadpoint>/vbapi_examplesdemonstrates |                                   |     |     |     | theusageof |
externaldata. Itprovidesutility methodsto convertaVB hashtableto amodel's
externaldata.
Exceptions
| Mostexceptionsthrownby |     | externaldatamethodsin |     | theVB | APIextend |
| ---------------------- | --- | --------------------- | --- | ----- | --------- |
IpfcXExternalDataError,whichisa
subclassof
IpfcXToolkitError.
| Anadditionalexceptionthrown |     | byexternaldatamethodsis |     |     |     |
| --------------------------- | --- | ----------------------- | --- | --- | --- |
IpfcXBadExternalData.Thisexception signalsan error accessingdata. For
example,externaldataaccessmighthavebeenterminatedorthe modelmight
| containstreamdatafromCreo |     | TOOLKIT.         |     |       |     |
| ------------------------- | --- | ---------------- | --- | ----- | --- |
| Thefollowingtablelists    |     | theseexceptions. |     |       |     |
| Exception                 |     |                  |     | Cause |     |
IpfcXExternalDataInvalidObject
Generatedwhenamodelorclassis
invalid.
IpfcXExternalDataClassOrSlotExists Generatedwhencreatingaclassorslot
andtheproposedclassorslotalready
exists.
IpfcXExternalDataNamesTooLong Generatedwhenaclassorslotnameis
toolong.
| IpfcXExternalDataSlotNotFound |     |     |     | Generatedwhenaspecifiedclassor |     |
| ----------------------------- | --- | --- | --- | ------------------------------ | --- |
slotdoesnotexist.
ExternalData 339

Exception Cause
IpfcXExternalDataEmptySlot Generatedwhentheslotyouare
attemptingtoreadisempty.
IpfcXExternalDataInvalidSlotName Generatedwhenaspecifiedslotname
isinvalid.
IpfcXBadGetExternalData Generatedwhenyoutrytoaccessan
incorrectdatatypeina
IpfcExternalDataobject.
340 Creo®ParametricVBAPIUser’sGuide

28
Windchill Connectivity APIs
Introduction.............................................................................................................342
AccessingaWindchillServerfromaCreoParametricSession...................................342
AccessingWorkspaces............................................................................................345
WorkflowtoRegisteraServer..................................................................................346
AliasedURL............................................................................................................347
ServerOperations...................................................................................................348
UtilityAPIs..............................................................................................................359
SampleBatchWorkflow...........................................................................................359
CreoParametrichasthecapabilityto be directlyconnectedto Windchillsolutions,
includingWindchillProjectLinkandPDMLinkservers. Thisaccess allowsusersto
manageandcontroltheproductdataseamlessly fromwithin Creo Parametric.
Thischapterlists theVB APIsthatsupport Windchillserversandserver
operationsina connectedCreo Parametricsession.
341

Introduction
Themethodsintroduced in thischapterprovidesupportfor thebasicWindchill
serveroperationsfrom within CreoParametric.Withthesemethods,operations
suchasregistering aWindchillserver,managing workspaces,andcheckin or
checkoutofobjectswillbe possibleviathe VBAPI.The capabilitiesofthese
APIsaresimilar tothe operationsavailablefromwithin theCreo Parametric
| client,with somerestrictions. |     |     |     |     |     |
| ----------------------------- | --- | --- | --- | --- | --- |
SomeoftheseAPIs aresupportedfrom anon-interactive,thatis,batchmode
applicationor asynchronousapplication.For moreinformationaboutbatch mode
| andasynchronousmode,refer | tothe       | chapteronAsynchronousmode. |        |      |        |
| ------------------------- | ----------- | -------------------------- | ------ | ---- | ------ |
| Accessing                 | a Windchill |                            | Server | from | a Creo |
| Parametric                | Session     |                            |        |      |        |
CreoParametricallowsyouto registerWindchillserversasaconnectionbetween
theWindchilldatabase andCreo Parametric.Although therepresented Windchill
databasecanbe from WindchillProjectLinkorWindchillPDMLinkalltypesof
| databasesarerepresented | in thesameway. |     |     |     |     |
| ----------------------- | -------------- | --- | --- | --- | --- |
You canusethefollowingidentifierswhenreferringto Windchillserversin the
VBAPI:
• CodebaseURL—This istheroot portionoftheURL thatisusedto connectto
aWindchillserver.For example<server.mycompany.com>/
Windchill.
• ServerAlias—Aserveraliasisused torefer to theserverafterithasbeen
registered.Thealiasis alsousedto construct pathstofilesin theserver
workspacesandcommonspaces.The serveraliasischosenby theuseror
applicationand itneednothaveanydirectrelationshipto the codebaseURL.
normalname,suchasmy_alias.
| Analiascan | beany       |        |             |     |        |
| ---------- | ----------- | ------ | ----------- | --- | ------ |
| Accessing  | Information | Before | Registering | a   | Server |
Tostartworking witha Windchillserver,youmustestablish aconnectionby
registeringtheserverin Creo Parametric.The methodsdescribed in thissection
allowyoutoconnect to aWindchillserverandaccess informationrelatedto the
server.
MethodsandPropertiesIntroduced:
• IpfcBaseSession.AuthenticateBrowser()
• IpfcBaseSession.GetServerLocation()
• IpfcServerLocation.Class
| 342 |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | ------------------------------- | --- | --- |

• IpfcServerLocation.Location
• IpfcServerLocation.Version
• IpfcServerLocation.ListContexts()
• IpfcServerLocation.CollectWorkspaces()
methodIpfcBaseSession.AuthenticateBrowser()
| Usethe |     |     |     |     |     | tosetthe |
| ------ | --- | --- | --- | --- | --- | -------- |
authenticationcontextusing avalid usernameandpassword.Asuccessful callto
thismethodallowsthe CreoParametricsessionto registerwith anyserverthat
acceptstheusername andpasswordcombination.Asuccessfulcallto thismethod
alsoensuresthatan authenticationdialogboxdoesnotappear duringthe
registrationprocess.You cancallthismethod anynumberoftimes tosetthe
authenticationcontextfor anynumberofWindchillservers,provided thatyou
registertheappropriateserversorserversimmediatelyaftersettingthecontext.
| ThepropertyIpfcServerLocation.Location |     |     |     |     | specifiesa |     |
| -------------------------------------- | --- | --- | --- | --- | ---------- | --- |
IpfcServerLocationobjectrepresentingthecodebase URLfor apossible
server.The servermaynothavebeenregisteredyet,butyoucan usethisobject
andthemethodsitcontains togather informationabout theserverpriorto
registration.
ThepropertyIpfcServerLocation.Class specifiesthe classof theserver
| orserverlocation.The |     | valuesare: |                         |     |     |     |
| -------------------- | --- | ---------- | ----------------------- | --- | --- | --- |
| • Windchill—Denotesa |     |            | WindchillPDMLinkserver. |     |     |     |
• ProjectLink—DenotesWindchillProjectLinktypeofservers.
| ThepropertyIpfcServerLocation.Version |     |     |     |     | specifies theversionof |     |
| ------------------------------------- | --- | --- | --- | --- | ---------------------- | --- |
Windchillthatisconfiguredon theserveror serverlocation,for example,9.0 or
| 10.0.This | methodacceptsthe |     | servercodebase | URLas | theinput. |     |
| --------- | ---------------- | --- | -------------- | ----- | --------- | --- |
Note
| IpfcServerLocation.Version |     |     |     | worksonly | forWindchillserversand |     |
| -------------------------- | --- | --- | --- | --------- | ---------------------- | --- |
throws theIpfcXToolkitUnsupported exception, iftheserveris nota
Windchillserver.
| ThemethodIpfcServerLocation.ListContexts() |     |     |     |     | givesalist | ofall |
| ------------------------------------------ | --- | --- | --- | --- | ---------- | ----- |
theavailablecontextsfor aspecifiedserver.Acontext isusedto associatea
| workspacewith                                   | aproduct,project,or |     | library. |     |     |            |
| ----------------------------------------------- | ------------------- | --- | -------- | --- | --- | ---------- |
| ThemethodIpfcServerLocation.CollectWorkspaces() |                     |     |          |     |     | returnsthe |
listofavailableworkspaces forthe specifiedserver.The workspaceobjects
| returnedcontainthe        |     | nameofeachworkspace |     | anditscontext. |     |     |
| ------------------------- | --- | ------------------- | --- | -------------- | --- | --- |
| WindchillConnectivityAPIs |     |                     |     |                |     | 343 |

| Registering | and Activating | a   | Server |     |
| ----------- | -------------- | --- | ------ | --- |
FromCreo Parametric2.0onward,thethe VBAPI methodscallthe same
underlyingAPIas CreoParametricto registerandunregisterservers. Hence,
registeringtheserversusing theVB APImethods issimilar toregistering the
serversusing theCreo Parametricuserinterface.Therefore, theserversregistered
bytheVB API areavailablein the CreoParametricServerRegistry.The servers
arealsoavailablein otherlocationsinthe CreoParametricuserinterface suchas,
| theFolderNavigatorandthe | embeddedbrowser. |     |     |     |
| ------------------------ | ---------------- | --- | --- | --- |
MethodsIntroduced:
• IpfcBaseSession.RegisterServer()
• IpfcServer.Activate()
• IpfcServer.Unregister()
| ThemethodIpfcBaseSession.RegisterServer() |     |     |     | registers the |
| ----------------------------------------- | --- | --- | --- | ------------- |
specifiedserverwith thecodebase URL.You canautomate theregistrationof
| serversin interactivemode.Topreregister |     | theserversusethestandard |     |     |
| --------------------------------------- | --- | ------------------------ | --- | --- |
config.fldsetup.Ifyoudo notwant theserversto be preregisteredin batch
mode,setthe environmentvariablePTC_WF_ROOT to anemptydirectorybefore
startingCreoParametric.
Asuccessfulcallto IpfcBaseSession.AuthenticateBrowser() witha
| validusername                    | andpasswordisessentialfor |     |               |               |
| -------------------------------- | ------------------------- | --- | ------------- | ------------- |
| IpfcBaseSession.RegisterServer() |                           |     | toregisterthe | serverwithout |
launchingtheauthenticationdialogbox. Registrationof theserverestablishesthe
serveralias.Youmustdesignatean existingworkspace to usewhen registeringthe
server.After theserverhasbeenregistered,youmaycreateanewworkspace.
ThemethodIpfcServer.Activate()sets thespecifiedserverasthe active
| serverin theCreo | Parametricsession. |     |     |     |
| ---------------- | ------------------ | --- | --- | --- |
ThemethodIpfcServer.Unregister() unregistersthespecifiedserver.
| Thisissimilar | to ServerRegistry▶▶Deletethrough |      | theuserinterface. |        |
| ------------- | -------------------------------- | ---- | ----------------- | ------ |
| Accessing     | Information                      | From | a Registered      | Server |
PropertiesIntroduced:
• IpfcServer.IsActive
• IpfcServer.Alias
• IpfcServer.Context
ThepropertyIpfcServer.IsActive specifies ifthe serveris active.
ThepropertyIpfcServer.Alias returnsthealiasof aserverifyou specify
thecodebaseURL.
| 344 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

ThepropertyIpfcServer.Context returnstheactivecontext ofthe active
server.
| Information | on Servers | in Session |     |     |     |
| ----------- | ---------- | ---------- | --- | --- | --- |
MethodsIntroduced:
• IpfcBaseSession.GetActiveServer()
• IpfcBaseSession.GetServerByAlias()
• IpfcBaseSession.GetServerByUrl()
• IpfcBaseSession.ListServers()
ThemethodIpfcBaseSession.GetActiveServer()
|     |     |     |     | returnsthe | active |
| --- | --- | --- | --- | ---------- | ------ |
serverhandle.
| ThemethodIpfcBaseSession.GetServerByAlias() |                        |     |                    | returnsthehandle |     |
| ------------------------------------------- | ---------------------- | --- | ------------------ | ---------------- | --- |
| tothe servermatching                        | thegivenserveralias,if |     | itexistsinsession. |                  |     |
ThemethodIpfcBaseSession.GetServerByUrl() returns thehandleto
theservermatchingthegiven serverURLandworkspace name,ifitexistsin
session.
ThemethodIpfcBaseSession.ListServers()
|                  |            |     | returnsalist | ofservers |     |
| ---------------- | ---------- | --- | ------------ | --------- | --- |
| registeredinthis | session.   |     |              |           |     |
| Accessing        | Workspaces |     |              |           |     |
Foreveryworkspace,a newdistinctstorage locationismaintainedin theuser’s
personalfolderonthe server(server-sideworkspace) andon theclient(client-side
workspacecache).Together,the server-side workspaceand theclient-side
| workspacecachemake | upthe workspace. |     |     |     |     |
| ------------------ | ---------------- | --- | --- | --- | --- |
MethodsandPropertiesIntroduced:
• CCpfcWorkspaceDefinition.Create()
• IpfcWorkspaceDefinition.WorkspaceName
• IpfcWorkspaceDefinition.WorkspaceContext
TheclassIpfcWorkspaceDefinition contains thenameand contextofthe
workspace.ThemethodIpfcServerLocation.CollectWorkspaces()
returnsan arrayof workspacedata. Workspace datais alsorequired forthe
| methodIpfcServer.CreateWorkspace()         |                  |     | tocreatea | workspacewith | a   |
| ------------------------------------------ | ---------------- | --- | --------- | ------------- | --- |
| givennameanda                              | specificcontext. |     |           |               |     |
| ThemethodCCpfcWorkspaceDefinition.Create() |                  |     |           | createsanew   |     |
workspacedefinitionobjectsuitablefor usewhencreatinganewworkspace on
theserver.
| WindchillConnectivityAPIs |     |     |     |     | 345 |
| ------------------------- | --- | --- | --- | --- | --- |

ThepropertyIpfcWorkspaceDefinition.WorkspaceName retrieves the
| nameofthe | workspace. |     |     |     |
| --------- | ---------- | --- | --- | --- |
ThepropertyIpfcWorkspaceDefinition.WorkspaceContext
| retrievesthecontext | ofthe workspace. |     |           |     |
| ------------------- | ---------------- | --- | --------- | --- |
| Creating            | and Modifying    | the | Workspace |     |
MethodsandPropertiesIntroduced:
• IpfcServer.CreateWorkspace()
• IpfcServer.ActiveWorkspace
• IpfcServerLocation.DeleteWorkspace()
ThemethodIpfcServer.CreateWorkspace() createsandactivates anew
workspace.
ThepropertyIpfcServer.ActiveWorkspace retrieves thenameof the
activeworkspace.
ThemethodIpfcServerLocation.DeleteWorkspace() deletesthe
specifiedworkspace.The methoddeletesthe workspaceonly ifthe following
conditionsaremet:
| • Theworkspace             | isnotthe                 | activeworkspace.      |                         |     |
| -------------------------- | ------------------------ | --------------------- | ----------------------- | --- |
| • Theworkspace             | doesnotcontain           | anycheckedoutobjects. |                         |     |
| Useoneof                   | thefollowingtechniquesto | deletean              | activeworkspace:        |     |
| • Maketherequiredworkspace |                          | inactiveusing         |                         |     |
| IpfcServer.ActiveWorkspace |                          |                       | withthe nameofsomeother |     |
workspaceandthencall
IpfcServerLocation.DeleteWorkspace().
• UnregistertheserverusingIpfcServer.Unregister() anddeletethe
workspace.
| Workflow    | to Register |         | a Server |           |
| ----------- | ----------- | ------- | -------- | --------- |
| To Register | a Server    | with an | Existing | Workspace |
Performthe followingsteps to registeraWindchillserverwith an existing
workspace:
| 1. Settheappropriateauthenticationcontext |     |     | using themethod |     |
| ----------------------------------------- | --- | --- | --------------- | --- |
IpfcBaseSession.AuthenticateBrowser()
witha validusername
andpassword.
| 346 |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | ------------------------------- | --- |

| 2. Lookupthe                            | listof workspacesusing |                        | themethod |                  |           |
| --------------------------------------- | ---------------------- | ---------------------- | --------- | ---------------- | --------- |
| IpfcServerLocation.CollectWorkspaces(). |                        |                        |           | Ifyoualreadyknow |           |
| thenameof                               | theworkspace           | ontheserver,thenignore |           | thisstep.        |           |
| 3. Registerthe                          | workspaceusing         | themethod              |           |                  |           |
| IpfcBaseSession.RegisterServer()        |                        |                        |           | with anexisting  | workspace |
| nameon                                  | theserver.             |                        |           |                  |           |
themethodIpfcServer.Activate().
4. Activatetheserverusing
| To Register | a Server | with | a New | Workspace |     |
| ----------- | -------- | ---- | ----- | --------- | --- |
Performthe followingsteps to registeraWindchillserverwith anewworkspace:
1. Perform steps1to 4in theprecedingsectionto registertheWindchillserver
| withan             | existing workspace.                     |         |     |     |           |
| ------------------ | --------------------------------------- | ------- | --- | --- | --------- |
| 2. Usethe          | methodIpfcServerLocation.ListContexts() |         |     |     | to choose |
| therequiredcontext | forthe                                  | server. |     |     |           |
3. Createanewworkspace withtherequired contextusing themethod
| IpfcServer.CreateWorkspace().This |         |     |     | methodautomatically | makes |
| --------------------------------- | ------- | --- | --- | ------------------- | ----- |
| thecreatedworkspace               | active. |     |     |                     |       |
Note
| Youcan  | createaworkspaceonly |     | aftertheserverisregistered. |     |     |
| ------- | -------------------- | --- | --------------------------- | --- | --- |
| Aliased | URL                  |     |                             |     |     |
AnaliasedURL servesas ahandleto theserverobjects.You canaccessthe server
objectsinthe commonspace(sharedfolders)and theworkspaceusing an aliased
URL.AnaliasedURL isaunique identifier forthe serverobjectand itsformatis
asfollows:
| • Objectinworkspace | hasaprefixwtws |     |     |     |     |
| ------------------- | -------------- | --- | --- | --- | --- |
wtws://<server_alias>/<workspace_name>/<object_server_name>
| where<object_server_name>includes |     |     |     | <object_ |     |
| --------------------------------- | --- | --- | --- | -------- | --- |
name>.<object_extension>
Forexample,
wtws://my_server/my_workspace/abcd.prt,
wtws://my_server/my_workspace/intf_file.igs
where
<server_alias>ismy_server
WindchillConnectivityAPIs 347

| <workspace_name>is              |     | my_workspace |     |
| ------------------------------- | --- | ------------ | --- |
| • Objectincommonspacehasaprefix |     | wtpub        |     |
wtpub://<server_alias>/<folder_location>/<object_server_name>
Forexample,
wtpub://my_server/path/to/cs_folder/abcd.prt
where
<server_alias>ismy_server
| <folder_location> |     | ispath/to/cs_folder |     |
| ----------------- | --- | ------------------- | --- |
Note
| ○ object_server_name |                         | mustbe         | in lowercase. |
| -------------------- | ----------------------- | -------------- | ------------- |
| ○ The                | APIsarecase-sensitiveto | thealiasedURL. |               |
○ <object_extension> shouldnotcontainCreo Parametricversions,
| for example,.1or.2,andso |            | on. |     |
| ------------------------ | ---------- | --- | --- |
| Server                   | Operations |     |     |
Afterregistering theWindchillserverwith Creo Parametric,youcan start
accessingthe dataon theWindchillservers.The Creo Parametricinteractionwith
Windchillserversleveragesthefollowinglocations:
• Commonspace(Sharedfolders)
| • Workspace | (Server-side | workspace)             |     |
| ----------- | ------------ | ---------------------- | --- |
| • Workspace | localcache   | (Client-sideworkspace) |     |
• CreoParametricsession
• Localdisk
Themethodsdescribed in thissectionenableyouto performthebasicserver
operations.The followingillustrationshowshowdataistransferredamong these
locations.
| 348 |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | ------------------------------- |

Save
MethodIntroduced:
• IpfcModel.Save()
ThemethodIpfcModel.Save() storesthe objectfromthe sessioninthe local
workspacecache,when aserverisactive.
Upload
Anupload transfersCreo Parametricfilesandanyotherdependencies fromthe
localworkspacecache to theserver-side workspace.
MethodsIntroduced:
• IpfcServer.UploadObjects()
• IpfcServer.UploadObjectsWithOptions()
• CCpfcUploadOptions.Create()
WindchillConnectivityAPIs 349

ThemethodIpfcServer.UploadObjects() uploads theobjectto the
workspace.Theobjectto be uploadedmustbe present inthe currentCreo
Parametricsession.Youmustsave theobjectto theworkspaceusing
IpfcModel.Save()IpfcBaseSession.ImportToCurrentWS()
beforeattemptingto upload it.
ThemethodIpfcServer.UploadObjectsWithOptions() uploads
objectstothe workspaceusing theoptions specifiedinthe
IpfcIpfcUploadOptions interface.Theseoptionsallow youtoupload the
entireworkspace,auto-resolvemissingreferences, andindicatethe targetfolder
locationfor thenewcontent duringthe upload.You mustsavethe objecttothe
workspaceusing IpfcModel.Save(),orimportitto theworkspace using
IpfcBaseSession.ImportToCurrentWS()before attemptingto upload
it.
Createthe IpfcUploadOptions objectusingthe method
CCpfcUploadOptions.Create().
Themethodsavailablefor settingtheupload optionsare describedin the
followingsection.
CheckIn
Afteryouhavefinishedworking onobjectsin yourworkspace,youcan sharethe
designchangeswithotherusers. The checkinoperationcopies theinformationand
filesassociated withall changedobjectsfrom theworkspaceto theWindchill
database.
MethodsandPropertiesIntroduced:
• IpfcServer.CheckinObjects()
• CCpfcCheckinOptions.Create()
• IpfcUploadBaseOptions.DefaultFolder
• IpfcUploadBaseOptions.NonDefaultFolderAssignments
• IpfcUploadBaseOptions.AutoresolveOption
• IpfcCheckinOptions.BaselineName
• IpfcCheckinOptions.BaselineNumber
• IpfcCheckinOptions.BaselineLocation
• IpfcCheckinOptions.BaselineLifecycle
• IpfcCheckinOptions.KeepCheckedout
ThemethodIpfcServer.CheckinObjects() checksin an objectintothe
database.Theobjectto be checkedin mustbe presentin thecurrentCreo
Parametricsession.Changesmade tothe objectarenotincludedunlessyousave
theobjectto theworkspaceusing themethod IpfcModel.Save()beforeyou
checkitin.
350 Creo®ParametricVBAPIUser’sGuide

IfyoupassNULLas thevalue oftheoptions parameter, thecheckinoperation is
similarto theAutoCheck-Inoption inCreo Parametric.Formoredetails onAuto
| Check-In,refer |                                    | tothe | onlinehelpfor |     | Creo | Parametric. |     |           |     |
| -------------- | ---------------------------------- | ----- | ------------- | --- | ---- | ----------- | --- | --------- | --- |
| Usethe         | methodCCpfcCheckinOptions.Create() |       |               |     |      |             |     | tocreatea | new |
IpfcCheckinOptionsobject.
Byusing an appropriatelyconstructedoptions argument, youcancontrolthe
checkinoperation.Usethe APIslisted aboveto accessandmodifythecheckin
| options.The |     | checkinoptions |     | areasfollows: |     |     |     |     |     |
| ----------- | --- | -------------- | --- | ------------- | --- | --- | --- | --- | --- |
• DefaultFolder—Specifiesthedefaultfolder locationonthe serverfor the
automaticcheckinoperation.
• NonDefaultFolderAssignment—Specifiesthefolderlocationon theserverto
|     | whichtheobjectswillbe |     |     | checkedin. |     |     |     |     |     |
| --- | --------------------- | --- | --- | ---------- | --- | --- | --- | --- | --- |
• AutoresolveOption—Specifiestheoption usedfor auto-resolvingmissing
|     | references.These                           |     | optionsaredefinedin |     |     | the |     |                       |          |
| --- | ------------------------------------------ | --- | ------------------- | --- | --- | --- | --- | --------------------- | -------- |
|     | EpfcServerAutoresolveOptionenumeratedtype, |     |                     |     |     |     |     | andareas              | follows: |
|     | ○ EpfcSERVER_DONT_AUTORESOLVE—Model        |     |                     |     |     |     |     | referencesmissingfrom |          |
theworkspace arenotautomaticallyresolved.This mayresultin aconflict
|     | uponcheckin.                            |     | Thisoption |     | isusedby      | default. |     |               |     |
| --- | --------------------------------------- | --- | ---------- | --- | ------------- | -------- | --- | ------------- | --- |
|     | ○ EpfcSERVER_AUTORESOLVE_IGNORE—Missing |     |            |     |               |          |     | referencesare |     |
|     | automaticallyresolvedby                 |     |            |     | ignoringthem. |          |     |               |     |
○ EpfcSERVER_AUTORESOLVE_UPDATE_IGNORE—Missing
referencesareautomaticallyresolvedby updating themin thedatabase and
|     | ignoringthem |     | ifnotfound. |     |     |     |     |     |     |
| --- | ------------ | --- | ----------- | --- | --- | --- | --- | --- | --- |
• Baseline—Specifiesthebaselineinformation forthe objectsuponcheckin.
|     | Thebaseline                       | informationfor |     |     | acheckin | operationisasfollows: |              |     |     |
| --- | --------------------------------- | -------------- | --- | --- | -------- | --------------------- | ------------ | --- | --- |
|     | ○ BaselineName—Specifiesthenameof |                |     |     |          | thebaseline.          |              |     |     |
|     | ○ BaselineNumber—Specifiesthe     |                |     |     |          | numberof              | thebaseline. |     |     |
Thedefaultformatfor thebaseline nameandbaseline numberisUsername
|     | +time(GMT)in                                 |     | milliseconds. |     |     |     |               |           |     |
| --- | -------------------------------------------- | --- | ------------- | --- | --- | --- | ------------- | --------- | --- |
|     | ○ BaselineLocation—Specifiesthelocationofthe |     |               |     |     |     |               | baseline. |     |
|     | ○ BaselineLifecycle—Specifiesthenameof       |     |               |     |     |     | thelifecycle. |           |     |
• KeepCheckedout—Ifthevaluespecifiedistrue, thenthecontents ofthe
selectedobjectare checkedintotheWindchillserverand automatically
|     | checkedoutagain |     | forfurther |     | modification. |     |     |     |     |
| --- | --------------- | --- | ---------- | --- | ------------- | --- | --- | --- | --- |
WindchillConnectivityAPIs 351

Retrieval
StandardVBAPI providesseveralmethodsthatarecapable ofretrieving models.
Whenusingthesemethods withWindchillservers,remember thatthesemethods
donotcheck outtheobjectto allowmodifications.
MethodsIntroduced:
• IpfcBaseSession.RetrieveModel()
• IpfcBaseSession.RetrieveModelWithOpts()
• IpfcBaseSession.OpenFile()
ThemethodsIpfcBaseSession.RetrieveModel(),
IpfcBaseSession.RetrieveModelWithOpts(),and
IpfcBaseSession.OpenFile()loadan objectintoasessiongivenits
nameandtype. Themethodssearch forthe objectinthe activeworkspace,the
localdirectory,and anyotherpathsspecifiedbythesearch_path configuration
option.
Checkout and Download
Tomodifyan objectfrom thecommonspace,youmustcheck outtheobject.The
processofcheckingoutcommunicatesyourintentionto modifyadesign tothe
Windchillserver. Theobjectin thedatabase islocked, sothatotheruserscan
obtainread-onlycopiesof theobject,andarepreventedfrom modifyingthe object
whileyouhavecheckeditout.
Checkoutis oftenaccompaniedbya downloadaction,wheretheobjectsare
broughtfrom theserver-side workspaceto thelocal workspacecache.In theVB
API,bothoperationsare coveredbythe samesetofmethods.
MethodsandPropertiesIntroduced:
• IpfcServer.CheckoutObjects()
• IpfcServer.CheckoutMultipleObjects()
• CCpfcCheckoutOptions.Create
• IpfcCheckoutOptions.Dependency
• IpfcCheckoutOptions.SelectedIncludes
• IpfcCheckoutOptions.IncludeInstances
• IpfcCheckoutOptions.Version
• IpfcCheckoutOptions.Download
• IpfcCheckoutOptions.Readonly
ThemethodIpfcServer.CheckoutObjects() checks outandoptionally
downloadstheobjectto theworkspace basedonthe configurationspecifications
oftheworkspace.The inputargumentsofthis methodareas follows:
352 Creo®ParametricVBAPIUser’sGuide

• Mdl—Specifiestheobjectto be checkedout.This isapplicableif themodel
hasalreadybeenretrievedwithoutcheckingitout.
| • File—Specifiesthetop-level |     | objectto | be checkedout. |     |     |
| ---------------------------- | --- | -------- | -------------- | --- | --- |
• Checkout—Thecheckoutflag. Ifyouspecify thevalueof thisargumentas
true,the selectedobjectis checkedout.Otherwise,the objectisdownloaded
withoutbeingcheckedout.The downloadactionenablesyouto bringread-
onlycopies ofobjectsintoyourworkspace.This allowsyouto examinethe
objectwithoutlockingit.
• Options—Specifiesthe checkoutoptions object.IfyoupassNULLasthe value
ofthis argument,, thenthedefaultCreo Parametriccheckoutrulesapply.Use
| themethodCCpfcCheckoutOptions.Create() |     |         |     | to createa | new |
| -------------------------------------- | --- | ------- | --- | ---------- | --- |
| IpfcCheckoutOptions                    |     | object. |     |            |     |
methodIpfcServer.CheckoutMultipleObjects()
| Usethe |     |     |     |     | to check |
| ------ | --- | --- | --- | --- | -------- |
outanddownloadmultiple objectsto theworkspacebasedon theconfiguration
specificationsoftheworkspace.This methodtakesthe sameinputargumentsas
listedabove,exceptfor MdlandFile.Instead ittakestheargumentFilesthat
| specifiesthesequence | ofthe | objectstocheck | outordownload. |     |     |
| -------------------- | ----- | -------------- | -------------- | --- | --- |
Note
| VB APImethodsdo | notsupport | AS_STOREDconfiguration. |     |     |     |
| --------------- | ---------- | ----------------------- | --- | --- | --- |
Byusing an appropriatelyconstructedoptions argumentin theabovefunctions,
youcancontrolthecheckoutoperation.Use theAPIslisted aboveto modifythe
| checkoutoptions. | Thecheckoutoptions | areasfollows: |     |     |     |
| ---------------- | ------------------ | ------------- | --- | --- | --- |
• Dependency—Specifiesthe dependencyrule usedwhile checkingout
dependentsofthe objectselectedforcheckout. Thetypesof dependencies
theEpfcServerDependency
| givenby                                                 |     |     | enumeratedtypeareasfollows: |                            |     |
| ------------------------------------------------------- | --- | --- | --------------------------- | -------------------------- | --- |
| ○ EpfcSERVER_DEPENDENCY_ALL—All                         |     |     |                             | theobjectsthataredependent |     |
| on theselectedobjectaredownloaded,thatis,theyareaddedto |     |     |                             |                            | the |
workspace.
| ○ EpfcSERVER_DEPENDENCY_REQUIRED—Allthe |     |     |     | objectsthatare |     |
| --------------------------------------- | --- | --- | --- | -------------- | --- |
requiredto successfullyretrieve theselectedobjectin theCADapplication
| aredownloaded,thatis,theyareaddedto                          |     |     | workspace. |                |         |
| ------------------------------------------------------------ | --- | --- | ---------- | -------------- | ------- |
| ○ EpfcSERVER_DEPENDENCY_NONE—None                            |     |     |            | ofthedependent | objects |
| from theselectedobjectaredownloaded,thatis,theyarenotaddedto |     |     |            |                |         |
workspace.
WindchillConnectivityAPIs 353

• IncludeInstances—Specifiestherule for includinginstances fromthe family
|     | tableduring                |     | checkout.The | typeof | instances | givenbythe                  |     |     |     |
| --- | -------------------------- | --- | ------------ | ------ | --------- | --------------------------- | --- | --- | --- |
|     | EpfcServerIncludeInstances |     |              |        |           | enumeratedtypeareasfollows: |     |     |     |
EpfcSERVER_INCLUDE_ALL—All
|     | ○                                 |     |                 |                |                     | theinstances        |                         | ofthe     | selectedobject |
| --- | --------------------------------- | --- | --------------- | -------------- | ------------------- | ------------------- | ----------------------- | --------- | -------------- |
|     | arechecked                        |     | out.            |                |                     |                     |                         |           |                |
|     | ○ EpfcSERVER_INCLUDE_SELECTED—The |     |                 |                |                     |                     | applicationcanselectthe |           |                |
|     | familytable                       |     | instancemembers |                | to beincludedduring |                     |                         | checkout. |                |
|     | ○ EpfcSERVER_INCLUDE_NONE—No      |     |                 |                |                     | additionalinstances |                         |           | fromthe        |
|     | familytable                       |     | areaddedto      | theobjectlist. |                     |                     |                         |           |                |
• SelectedIncludes—Specifiesthe sequenceofURLs to theselectedinstances,if
|     | IncludeInstancesis |     | oftypeEpfcSERVER_INCLUDE_SELECTED. |     |     |     |     |     |     |
| --- | ------------------ | --- | ---------------------------------- | --- | --- | --- | --- | --- | --- |
• Version—Specifiesthe versionof thecheckedout object.Ifthis valueis setto
NULL,the
|     |     | objectischeckedout |     |     | accordingtothe |     | currentworkspace |     |     |
| --- | --- | ------------------ | --- | --- | -------------- | --- | ---------------- | --- | --- |
configuration.
• Download—Specifiesthecheckouttypeasdownload orlink. Thevalue
downloadspecifiesthatthe objectcontentisdownloaded andcheckedout,
whilelinkspecifiesthatonly themetadataisdownloaded andcheckedout.
• Readonly—Specifiesthecheckouttypeasaread-only checkout. Thisoption is
|     | applicableonly |     | ifthecheckouttypeislink. |     |     |     |     |     |     |
| --- | -------------- | --- | ------------------------ | --- | --- | --- | --- | --- | --- |
Thefollowingtruthtableexplains thedependencies ofthedifferent controlfactors
inthe methodIpfcServer.CheckoutObjects() andtheeffect ofdifferent
| combinationson   |     |     | theend result. |     |                |     |     |        |     |
| ---------------- | --- | --- | -------------- | --- | -------------- | --- | --- | ------ | --- |
| Argumentcheckout |     |     | IpfcCheckoutOp |     | IpfcCheckoutOp |     |     | Result |     |
| in               |     |     | tions.Download |     | tions.Readonly |     |     |        |     |
IpfcServer.Check
outObjects()
| true |     |     | true |     | NA  |     |     | Objectischeckedoutand |     |
| ---- | --- | --- | ---- | --- | --- | --- | --- | --------------------- | --- |
itscontentis
downloaded.
| true |     |     |       |     | NA  |     |     |                       |     |
| ---- | --- | --- | ----- | --- | --- | --- | --- | --------------------- | --- |
|      |     |     | false |     |     |     |     | Objectischeckedoutbut |     |
contentisnot
downloaded.
| false |     |     | NA  |     | true |     |     | Objectisdownloaded |     |
| ----- | --- | --- | --- | --- | ---- | --- | --- | ------------------ | --- |
withoutcheckoutandas
read-only.
| false |          |     | NA  |     | false |     |     | Notsupported |     |
| ----- | -------- | --- | --- | --- | ----- | --- | --- | ------------ | --- |
| Undo  | Checkout |     |     |     |       |     |     |              |     |
MethodIntroduced:
• IpfcServer.UndoCheckout()
| 354 |     |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | --- | --- | --- | ------------------------------- | --- | --- |

Usethe methodIpfcServer.UndoCheckout()to undoa checkoutof the
specifiedobject.Whenyouundo acheckout, thechangesthatyouhavemadeto
thecontentandmetadataof theobjectarediscardedandthecontent,as storedin
theserver, isdownloaded tothe workspace.Thismethod isapplicableonly forthe
modelin theactive CreoParametricsession.
Import and Export
TheVB APIprovidesyouwith thecapabilityof transferringspecifiedobjectsto
andfroma workspace.Importand exportoperationsmusttakeplacein a session
withno models.An importoperationtransfers afilefromthe localdiskto the
workspace.
MethodsandPropertiesIntroduced:
• IpfcBaseSession.ExportFromCurrentWS()
• IpfcBaseSession.ImportToCurrentWS()
• IpfcWSImportExportMessage.Description
• IpfcWSImportExportMessage.FileName
• IpfcWSImportExportMessage.MessageType
• IpfcWSImportExportMessage.Resolution
• IpfcWSImportExportMessage.Succeeded
• IpfcBaseSession.SetWSExportOptions()
• CCpfcWSExportOptions.Create()
• IpfcWSExportOptions.IncludeSecondaryContent
ThemethodIpfcBaseSession.ExportFromCurrentWS() exports the
specifiedobjectsfrom thecurrentworkspace toa diskin alinkedsessionofCreo
Parametric.
ThemethodIpfcBaseSession.ImportToCurrentWS() importsthe
specifiedobjectsfrom adisk tothe currentworkspace in alinkedsessionofCreo
Parametric.
BothIpfcBaseSession.ExportFromCurrentWS() and
IpfcBaseSession.ImportToCurrentWS()allow youtospecify a
dependencycriterionto processthefollowingitems:
• Allexternaldependencies
• Onlyrequired dependencies
• Noexternaldependencies
BothIpfcBaseSession.ExportFromCurrentWS() and
IpfcBaseSession.ImportToCurrentWS()return themessages
generatedduringtheexport orimportoperationin theform ofthe
WindchillConnectivityAPIs 355

IpfcWSImportExportMessagesobject.Use theAPIs listedaboveto access
| thecontentsof                              |     | amessage.Themessagespecifiedby |     |     |     |     |     | the             |     |
| ------------------------------------------ | --- | ------------------------------ | --- | --- | --- | --- | --- | --------------- | --- |
| IpfcWSImportExportMessageobjectcontainsthe |     |                                |     |     |     |     |     | followingitems: |     |
• Description—Specifiesthe descriptionof theproblem orthemessage
information.
| •   | FileName—Specifiestheobjectnameor |     |     |     |     |     | thename | oftheobjectpath. |     |
| --- | --------------------------------- | --- | --- | --- | --- | --- | ------- | ---------------- | --- |
• MessageType—Specifiesthe severityofthe messagein theformof the
|     | EpfcWSImportExportMessageTypeenumerated |                 |     |     |     |     |                 |     | type.The severityis |
| --- | --------------------------------------- | --------------- | --- | --- | --- | --- | --------------- | --- | ------------------- |
|     | oneofthe                                | followingtypes: |     |     |     |     |                 |     |                     |
|     | ○ EpfcWSIMPEX_MSG_INFO—Specifies        |                 |     |     |     |     | aninformational |     | typeof              |
message.
|     | ○ EpfcWSIMPEX_MSG_WARNING—Specifies  |                     |     |     |                     |     |     | alowseverity       | problemthat |
| --- | ------------------------------------ | ------------------- | --- | --- | ------------------- | --- | --- | ------------------ | ----------- |
|     | canbe                                | resolvedaccordingto |     |     | theconfiguredrules. |     |     |                    |             |
|     | ○ EpfcWSIMPEX_MSG_CONFLICT—Specifies |                     |     |     |                     |     |     | aconflictthatcanbe |             |
overridden.
|     | ○ EpfcWSIMPEX_MSG_ERROR—Specifies |     |                                      |     |     |     |     | aconflictthatcannotbe |              |
| --- | --------------------------------- | --- | ------------------------------------ | --- | --- | --- | --- | --------------------- | ------------ |
|     | overriddenora                     |     | seriousproblemthatpreventsprocessing |     |     |     |     |                       | ofan object. |
• Resolution—Specifiestheresolutionapplied toresolve aconflictthatcan
|     | beoverridden.This |     | isapplicable |     | whenthemessageis |     |     | ofthe | type |
| --- | ----------------- | --- | ------------ | --- | ---------------- | --- | --- | ----- | ---- |
EpfcWSIMPEX_MSG_CONFLICT.
• Succeeded—Determineswhetherthe resolutionsucceededornot.This is
|     | applicablewhenthe |     | messageisof |     | thetypeEpfcWSIMPEX_MSG_ |     |     |     |     |
| --- | ----------------- | --- | ----------- | --- | ----------------------- | --- | --- | --- | --- |
CONFLICT.
| ThemethodIpfcBaseSession.SetWSExportOptions() |     |     |     |     |     |     |     |     | setstheexport |
| --------------------------------------------- | --- | --- | --- | --- | --- | --- | --- | --- | ------------- |
optionsusedwhile exportingtheobjectsfrom aworkspacein theform ofthe
| IpfcWSExportOptions |     |     |     | object.Createthis |     |     | objectusingthe |     | method |
| ------------------- | --- | --- | --- | ----------------- | --- | --- | -------------- | --- | ------ |
CCpfcWSExportOptions.Create(). Theexport optionsare asfollows:
• IncludeSecondaryContent—Indicateswhetherornot toinclude secondary
contentwhileexportingthe primaryCreoParametric modelfiles. Usethe
propertyIpfcWSExportOptions.IncludeSecondaryContent to
setthisoption.
File Copy
| TheVB             | APIprovidesyouwith |     |          | thecapabilityof |     |                       | copyingafilefromthe |     |     |
| ----------------- | ------------------ | --- | -------- | --------------- | --- | --------------------- | ------------------- | --- | --- |
| workspaceortarget |                    |     | folderto | alocationon     |     | thediskandvice-versa. |                     |     |     |
MethodsIntroduced:
• IpfcBaseSession.CopyFileToWS()
| 356 |     |     |     |     |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | --- | --- | --- | --- | ------------------------------- | --- |

• IpfcBaseSession.CopyFileFromWS()
methodIpfcBaseSession.CopyFileToWS()
| Usethe |     |     |     |     | to copyafilefrom |     |
| ------ | --- | --- | --- | --- | ---------------- | --- |
thediskto theworkspace.The filecanoptionallybe addedassecondary content
toa givenworkspacefile. Iftheviewablefileisaddedas secondarycontent,a
dependencyis createdbetweenthe CreoParametricmodelandthe viewablefile.
| Usethe           | methodIpfcBaseSession.CopyFileFromWS() |                |       |     | to copya | file |
| ---------------- | -------------------------------------- | -------------- | ----- | --- | -------- | ---- |
| fromtheworkspace |                                        | to alocationon | disk. |     |          |      |
Whenimportingor exportingCreoParametricmodels, PTCrecommendsthatyou
| usemethods | IpfcBaseSession.ImportToCurrentWS()and |     |     |     |     |     |
| ---------- | -------------------------------------- | --- | --- | --- | --- | --- |
IpfcBaseSession.ExportFromCurrentWS(),respectively,to perform
theimportorexport operation.Methods thatcopy individualfilesdo nottraverse
CreoParametricmodeldependencies,andtherefore do notcopya fullyretrievable
setof modelsatthesametime.
| Additionally, | onlythe | methods |     |     |     |     |
| ------------- | ------- | ------- | --- | --- | --- | --- |
IpfcBaseSession.ImportToCurrentWS()and
| IpfcBaseSession.ExportFromCurrentWS()providefull |                   |     |     |     | metadata       |     |
| ------------------------------------------------ | ----------------- | --- | --- | --- | -------------- | --- |
| exchangeand                                      | support.Thatmeans |     |     |     |                |     |
| IpfcBaseSession.ImportToCurrentWS()can           |                   |     |     |     | communicateall | the |
CreoParametricdesignated parameters,dependencies,and familytable
| informationto                            | aPDM | systemwhile |     |     |                    |     |
| ---------------------------------------- | ---- | ----------- | --- | --- | ------------------ | --- |
| IpfcBaseSession.ExportFromCurrentWS()can |      |             |     |     | updateexportedCreo |     |
Parametricdatawith PDM systemchangesto designatedand systemparameters,
dependencies,andfamilytable information.HencePTCrecommends theuseof
| IpfcBaseSession.CopyFileToWS()   |     |     |     | and       |              |     |
| -------------------------------- | --- | --- | --- | --------- | ------------ | --- |
| IpfcBaseSession.CopyFileFromWS() |     |     |     | toprocess | onlynon-Creo |     |
Parametricfiles.
| Server | Object | Status |     |     |     |     |
| ------ | ------ | ------ | --- | --- | --- | --- |
MethodsIntroduced:
• IpfcServer.IsObjectCheckedOut()
• IpfcServer.IsObjectModified()
Themethodsdescribed in thissectionverifythe currentstatusoftheobjectin the
| workspace.ThemethodIpfcServer.IsObjectCheckedOut() |     |     |     |     |     | specifies |
| -------------------------------------------------- | --- | --- | --- | --- | --- | --------- |
whetherthe objectischeckedout formodification.The valuetrueindicatesthat
| thespecifiedobjectischeckedout |     |                | tothe activeworkspace. |     |     |     |
| ------------------------------ | --- | -------------- | ---------------------- | --- | --- | --- |
| Thevaluefalse                  |     | indicatesoneof | thefollowingstatuses:  |     |     |     |
| • Thespecifiedobjectis         |     | notcheckedout  |                        |     |     |     |
• Thespecifiedobjectis onlyuploaded to theworkspace,butwas neverchecked
in
| WindchillConnectivityAPIs |     |     |     |     |     | 357 |
| ------------------------- | --- | --- | --- | --- | --- | --- |

• Thespecifiedobjectis onlysavedto thelocal workspacecache,butwas never
uploaded
ThemethodIpfcServer.IsObjectModified() specifieswhetherthe
objecthasbeenmodifiedin theworkspace.This methodreturnstrueif the
| objectwas | modifiedlocally. |     |     |     |     |     |
| --------- | ---------------- | --- | --- | --- | --- | --- |
| Delete    | Objects          |     |     |     |     |     |
MethodIntroduced:
• IpfcServer.RemoveObjects()
ThemethodIpfcServer.RemoveObjects() deletesthe arrayofobjects
fromtheworkspace. Whenpassedwiththe ModelNamesarrayasNULL, this
| methodremovesallthe |     |        | objectsinthe |        | activeworkspace. |     |
| ------------------- | --- | ------ | ------------ | ------ | ---------------- | --- |
| Conflicts           |     | During |              | Server | Operations       |     |
Anexceptionis providedto capturetheerror conditionwhileperforming the
| followingserveroperationsusing |     |     |     |     | thespecifiedAPIs:           |     |
| ------------------------------ | --- | --- | --- | --- | --------------------------- | --- |
| Operation                      |     |     |     |     | API                         |     |
| Checkinanobjectorworkspace     |     |     |     |     | IpfcServer.CheckinObjects() |     |
IpfcServer.CheckoutObjects()
Checkoutanobject
| Undocheckoutofanobject |     |     |     |     | IpfcServer.UndoCheckout()         |     |
| ---------------------- | --- | --- | --- | --- | --------------------------------- | --- |
| Uploadobject           |     |     |     |     | IpfcServer.UploadObjects()        |     |
| Downloadobject         |     |     |     |     | IpfcServer.CheckoutObjects()(with |     |
downloadastrue)
| Deleteworkspace |     |     |     |     | IpfcServerLocation.DeleteWork |     |
| --------------- | --- | --- | --- | --- | ----------------------------- | --- |
space()
| Removeobject   |     |         |     |                                       | IpfcServer.RemoveObjects() |     |
| -------------- | --- | ------- | --- | ------------------------------------- | -------------------------- | --- |
| TheseAPIsthrow |     | acommon |     | exceptionIpfcXToolkitCheckoutConflict |                            |     |
ifan erroris encounteredduringserveroperations.Theexception descriptionwill
includethedetailsof theerror condition.Thisdescriptionissimilar tothe
descriptiondisplayedby theCreo ParametricHTMLuserinterfacein theconflict
report.
| Example         | Code: |     | Server                             | APIs |     |          |
| --------------- | ----- | --- | ---------------------------------- | ---- | --- | -------- |
| Thesamplecodein |       |     | thefilepfcServerExamples.vblocated |      |     | at<creo_ |
vbapi_loadpoint>/vbapi_examplesdemonstrates theimplementationof
| theWindchillserverAPIs |     |     |     | describedin | theprevioussections. |     |
| ---------------------- | --- | --- | --- | ----------- | -------------------- | --- |
358 Creo®ParametricVBAPIUser’sGuide

| Utility | APIs |     |     |     |     |     |
| ------- | ---- | --- | --- | --- | --- | --- |
Themethodsspecifiedin thissectionenableyouto obtainthehandleto theserver
objectstoaccess them.Thehandlemaybethe aliasedURLor themodelnameof
thehttpURL.Theseutilitiesenablethe conversionofonetypeofhandleto
another.
MethodsIntroduced:
• IpfcServer.GetAliasedUrl()
• IpfcBaseSession.GetModelNameFromAliasedUrl()
• IpfcBaseSession.GetAliasFromAliasedUrl()
• IpfcBaseSession.GetUrlFromAliasedUrl()
| ThemethodIpfcServer.GetAliasedUrl() |     |     |     | enablesyouto | searchfor | a   |
| ----------------------------------- | --- | --- | --- | ------------ | --------- | --- |
serverobjectby itsname.Specifythecomplete filenameof theobjectasthe input,
forexample,test_part.prt.The methodreturnsthe aliasedURLfor amodel
ontheserver.For moreinformationregarding thealiasedURL,refer to thesection
AliasedURL onpage347. Duringthesearch operation,the workspacetakes
| precedenceover | thesharedspace. |     |     |     |     |     |
| -------------- | --------------- | --- | --- | --- | --- | --- |
You canalsousethis methodto searchforfilesthatarenotin the CreoParametric
| format.For | example,my_text.txt, |     | intf_file.stp,and |     | soon. |     |
| ---------- | -------------------- | --- | ----------------- | --- | ----- | --- |
ThemethodIpfcBaseSession.GetModelNameFromAliasedUrl()
returnsthename oftheobjectfrom thegivenaliased URLon theserver.
| ThemethodIpfcBaseSession.GetUrlFromAliasedUrl()    |                  |        |              |         | convertsan |     |
| -------------------------------------------------- | ---------------- | ------ | ------------ | ------- | ---------- | --- |
| aliasedURLto                                       | astandardURL     | forthe | objectsonthe | server. |            |     |
| Forexample,wtws://my_alias/CreoParametric/abcd.prt |                  |        |              |         | is         |     |
| convertedtoan                                      | appropriateURLon |        | theserveras  |         |            |     |
<server.mycompany.com>/Windchill.
| ThemethodIpfcBaseSession.GetAliasFromAliasedUrl() |       |          |     |     | returns |     |
| ------------------------------------------------- | ----- | -------- | --- | --- | ------- | --- |
| theserveraliasfromaliased                         |       | URL.     |     |     |         |     |
| Sample                                            | Batch | Workflow |     |     |         |     |
Atypicalworkflow usingthe WindchillAPIs foran asynchronousnon-graphical
| applicationisas | follows:                  |     |        |     |     |     |
| --------------- | ------------------------- | --- | ------ | --- | --- | --- |
| 1. StartaCreo   | Parametricsessionusingthe |     | method |     |     |     |
CCpfcAsyncConnection.Connect().
| 2. Authenticatethebrowserusing |     |     | themethod |     |     |     |
| ------------------------------ | --- | --- | --------- | --- | --- | --- |
IpfcBaseSession.AuthenticateBrowser().
| 3. Registerthe | serverwith | thenewworkspace | using | themethod |     |     |
| -------------- | ---------- | --------------- | ----- | --------- | --- | --- |
IpfcBaseSession.RegisterServer().
| WindchillConnectivityAPIs |     |     |     |     |     | 359 |
| ------------------------- | --- | --- | --- | --- | --- | --- |

| 4. Activatetheserverusing | themethodIpfcServer.Activate(). |     |     |
| ------------------------- | ------------------------------- | --- | --- |
5. Checkoutand retrievethe modelfrom thevault URLusing themethod
IpfcServer.CheckoutObjects() followedby
IpfcBaseSession.RetrieveModel().
| 6. Modifythe | modelaccordingto | theapplicationlogic. |     |
| ------------ | ---------------- | -------------------- | --- |
7. Savethe modeltothe workspaceusing themethodIpfcModel.Save().
| 8. Checkin | themodifiedmodelbackto | theserverusing | themethod |
| ---------- | ---------------------- | -------------- | --------- |
pfcServer.Server.CheckinObjects().
9. Afterprocessing all models,unregisterfrom theserverusing themethod
IpfcServer.Unregister().
10.Deletetheworkspace using
IpfcServerLocation.DeleteWorkspace().
11. StopCreo Parametricusing themethodIpfcAsyncConnection.End().
| 360 |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | ------------------------------- | --- |

29
Technical Summary of Changes
for Creo 9.0.0.0
NewFunctions........................................................................................................362
MiscellaneousTechnicalChanges............................................................................362
FullVersionofCreo®ParametricVBAPIReleaseNotes............................................363
Thischapterdescribes thecriticalandmiscellaneous technicalchangesin Creo
Parametric9.0.0.0andVB API.It alsolists thenewandsuperseded functions for
thisrelease.
361

| New | Functions |     |     |     |     |     |
| --- | --------- | --- | --- | --- | --- | --- |
Thissectiondescribesnew functionsfor VBAPI forCreo Parametric9.0.0.0.
| Action       | Listeners |     |             |     |     |     |
| ------------ | --------- | --- | ----------- | --- | --- | --- |
| New Function |           |     | Description |     |     |     |
IpfcSelectionBufferListener. Thismethod iscalledaftertheselection
| OnAfterSelBufferChange() |     |     | bufferis    | changed.Thismethod | is  |     |
| ------------------------ | --- | --- | ----------- | ------------------ | --- | --- |
|                          |     |     | availableby | calling themethod  |     |     |
IpfcActionSource.AddAction
|     |     |     | ListenerWithTypewith |        | thevalue |     |
| --- | --- | --- | -------------------- | ------ | -------- | --- |
|     |     |     | ofthenotify          | typeas |          |     |
EpfcSELBUFFER_CHANGE_POST.
| Miscellaneous |     | Technical |     | Changes |     |     |
| ------------- | --- | --------- | --- | ------- | --- | --- |
ThefollowingchangesinCreo 9.0.0.0 canaffectthe functionalbehavior ofVB
API.PTCdoesnot anticipatethatthesechangescausecriticalissueswith existing
| VBAPI applications.         |            |          |           |                       |     |     |
| --------------------------- | ---------- | -------- | --------- | --------------------- | --- | --- |
| Support                     | for Spline | Surfaces |           | with 2 Derivatives    |     |     |
| ThevalueEpfcSURFACE_SPL2DER |            |          | isaddedto | theenumerateddatatype |     |     |
EpfcSurfaceType,whichprovidessupportfor spline surfaceswith 2
derivatives.
| Support | for Datum | Curve | End | in Attachment |     |     |
| ------- | --------- | ----- | --- | ------------- | --- | --- |
Selection
| ThevalueEpfcITEM_CRV_START |     |     | andEpfcITEM_CRV_END |     | isaddedto | the |
| -------------------------- | --- | --- | ------------------- | --- | --------- | --- |
enumerateddatatypeEpfcModelItemType, whichsupports selectionof
specificdimensionattachment.
| Obsolete                          | Function |     |             |     |     |     |
| --------------------------------- | -------- | --- | ----------- | --- | --- | --- |
| InCreo Parametric9.0.0.0andlater, |          |     | theproperty |     |     |     |
IpfcModelDescriptor.Hostis obsoletedandwillnothaveanysupport.
| 362 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |     |
| --- | --- | --- | --- | ------------------------------- | --- | --- |

| Full Version  | of Creo® | Parametric | VB API |
| ------------- | -------- | ---------- | ------ |
| Release Notes |          |            |        |
Toseea fullversion oftheCreo®Parametric VBAPI ReleaseNotes,visitthe
pageCreo®ParametricVB APIRelease Notes.The fullversion contains
| informationfromall | thepast releasenotesforVB | API. |     |
| ------------------ | ------------------------- | ---- | --- |
TechnicalSummaryofChangesforCreo9.0.0.0 363

30
Sample Applications
InstallingtheVBAPI................................................................................................365
SampleApplications................................................................................................365
Thisappendix liststhe sampleapplicationsprovided withthe VBAPI.
364 Creo®ParametricVBAPIUser’sGuide

| Installing  | the         | VB        | API |        |                     |     |     |
| ----------- | ----------- | --------- | --- | ------ | ------------------- | --- | --- |
| TheVB APIis | availableon | thesameCD |     | asCreo | Parametric.WhenCreo |     |     |
Parametricisinstalledusing PTC.SetUp,oneofthe optionalcomponentsis API
Toolkits.This includesCreo TOOLKIT,J-Link,Web.Link, andVisualBasic
API.
IfyouselectVisualBasicAPI,adirectorycalledvbapi iscreated underthe Creo
ParametricloadpointandtheVB APIis automaticallyinstalledin this directory.
Thisdirectorycontainsall thelibraries, exampleapplications,anddocumentation
| specificto theVB | API.         |     |     |     |     |     |     |
| ---------------- | ------------ | --- | --- | --- | --- | --- | --- |
| Sample           | Applications |     |     |     |     |     |     |
TheVB APIsampleapplicationsareavailablein thedirectoriesvbapi_
| examplesandvbapi_applsunderthe |     |     |     | rootdirectoryvbapi. |     |     |     |
| ------------------------------ | --- | --- | --- | ------------------- | --- | --- | --- |
VBAPIExamples
| Location                      |     |     |     | MainClasses                   |     |     |     |
| ----------------------------- | --- | --- | --- | ----------------------------- | --- | --- | --- |
| <creo_vbapi_loadpoint>/vbapi_ |     |     |     | formConnectionandformExamples |     |     |     |
examples
Thevbapi_examplesdirectoryis acollectionof alltheVB.NETexample
sourcefilespresent in theVBAPI User’s Guide.Allthe examplesourcefilesare
alsoavailablealongwitha VB.NETsolutioncalledVB APIExamples.sln
| andaprojectfilecalledVB |                  | APIExamples.vbproj |                   |                     |     | asa single |     |
| ----------------------- | ---------------- | ------------------ | ----------------- | ------------------- | --- | ---------- | --- |
| VBAPIExamples.zipfilein |                  |                    | thesamedirectory. |                     |     |            |     |
| Setupandrunthe          | examplesusingthe |                    |                   | followingprocedure: |     |            |     |
1. SetthePRO_COMM_MSG_EXE environmentvariableto thefull pathofthe
executablepro_comm_msg.exefor yourapplicationto communicatewith
is<creo_
| CreoParametric.Typically,the             |     |     | pathto                          | theexecutable        |     |            |        |
| ---------------------------------------- | --- | --- | ------------------------------- | -------------------- | --- | ---------- | ------ |
| loadpoint>\<datecode>\Common             |     |     |                                 | Files\<machinetype>\ |     |            |        |
| obj\pro_comm_msg.exe,wheremachine_typeis |     |     |                                 |                      |     | i486_ntfor | 32-bit |
| Windowsandx86_win64                      |     |     | for 64-bitWindowsinstallations. |                      |     |            |        |
2. Registerthe COMserverbyrunning thevb_api_register.bat file
locatedat<creo_loadpoint>/Parametric/bin.
3. UnziptheVBAPIExamples.zip filein alocalfolder onyour machineand
| openthe                  | VB APIExamples.slnsolutionin |                 |               |                  | Microsoft | VisualStudio. |     |
| ------------------------ | ---------------------------- | --------------- | ------------- | ---------------- | --------- | ------------- | --- |
| 4. SettheCOMreferencefor |                              |                 | yourprojectto | CreoParametricVB |           |               | API |
| TypeLibraryfor           |                              | CreoParametric. |               |                  |           |               |     |
| SampleApplications       |                              |                 |               |                  |           |               | 365 |

5. Buildthe solutionandexecute theVB APIExamples.execreated inyour
| localfolder.The | VBAPIExamples-Connectionformas | shownin the |
| --------------- | ------------------------------ | ----------- |
| followingfigure | isloaded.                      |             |
6. ClickStartto startanewCreo Parametricsessionin thesimpleasynchronous
mode.You mustspecifythe CreoParametricworking directoryandexecutable
pathbeforeattemptingto startanewsession. You canalso connectto an
existingsessionin thesimpleasynchronous modebyclickingConnect. Click
BatchModeto startanewCreo Parametricsessionin thebatchmode,or click
| 366 | Creo®ParametricVBAPIUser’sGuide |     |
| --- | ------------------------------- | --- |

FullAsyncto startanewCreo Parametricsessioninthe fullasynchronous
| mode.Refertothe |     | chapterAsynchronousModeVB |     |     | APIFundamentals: |     |     |
| --------------- | --- | ------------------------- | --- | --- | ---------------- | --- | --- |
ControllingCreo Parametriconpage31 formoreinformationon themodes of
communication.
7. Onceyouare connectedto aCreo Parametricsession,the VB APIExamples
formis loaded.Youcan executeall theexamplesavailablein thevbapi_
| examplesdirectoryfromthis           |     |                | form. |           |     |     |     |
| ----------------------------------- | --- | -------------- | ----- | --------- | --- | --- | --- |
| Parameters                          |     | and Dimensions |       |           |     |     |     |
| Location                            |     |                |       | MainClass |     |     |     |
| <creo_vbapi_loadpoint>/vbapi_appls/ |     |                |       | FormPD    |     |     |     |
vbparam
Theparametersanddimensionsexampleis anasynchronous modeVB.NET
applicationthatallowsyouto accessand modifythe parametersand dimensions
ofaCreo Parametricmodel.All theVB sourcefilesfor thisapplicationare
| availablealongwith                   |     | aVB.NETsolution |     | called  |                   |     |     |
| ------------------------------------ | --- | --------------- | --- | ------- | ----------------- | --- | --- |
| ParameterAndDimension.slnandaproject |     |                 |     |         | filecalled        |     |     |
| ParameterAndDimension.vbprojas       |     |                 |     | asingle | VBParam.zipfilein |     | the |
vbparamdirectory.
| Setupandrunthis |     | applicationusing |     | thefollowingprocedure: |     |     |     |
| --------------- | --- | ---------------- | --- | ---------------------- | --- | --- | --- |
1. SetthePRO_COMM_MSG_EXE environmentvariableto thefull pathofthe
executablepro_comm_msg.exefor yourapplicationto communicatewith
| CreoParametric.Typically,the             |     |     | pathto                          | theexecutable        | is<creo_   |     |        |
| ---------------------------------------- | --- | --- | ------------------------------- | -------------------- | ---------- | --- | ------ |
| loadpoint>\<datecode>\Common             |     |     |                                 | Files\<machinetype>\ |            |     |        |
| obj\pro_comm_msg.exe,wheremachine_typeis |     |     |                                 |                      | i486_ntfor |     | 32-bit |
| Windowsandx86_win64                      |     |     | for 64-bitWindowsinstallations. |                      |            |     |        |
2. Registerthe COMserverbyrunning thevb_api_register.bat file
locatedat<creo_loadpoint>/Parametric/bin
3. UnziptheVBParam.zip fileina localfolderon yourmachineandopen the
| ParameterAndDimension.slnsolutionin |     |     |     |     | Microsoft | VisualStudio. |     |
| ----------------------------------- | --- | --- | --- | --- | --------- | ------------- | --- |
4. SettheCOMreferencefor yourprojecttoCreoParametricVB API Type
| Libraryfor  |                    | CreoParametric. |                              |     |     |     |     |
| ----------- | ------------------ | --------------- | ---------------------------- | --- | --- | --- | --- |
| 5. Buildthe | solutionandexecute |                 | theParameterAndDimension.exe |     |     |     |     |
createdin yourlocal folder. TheParametersandDimensions formis loaded.
6. StartCreoParametricand opena PART modelcontainingparametersand
dimensions.
7. Clickthe Connectbutton in theform toconnect to theactiveCreo Parametric
sessioninthe simpleasynchronousmode.Click theAddbutton to connectin
| SampleApplications |     |     |     |     |     |     | 367 |
| ------------------ | --- | --- | --- | --- | --- | --- | --- |

thefull asynchronousmode,whereinanewPDMenu menugetsadded tothe
| menubarin | theCreo | Parametricuserinterface. |     |     |     |     |
| --------- | ------- | ------------------------ | --- | --- | --- | --- |
You canperformthesamesetof operationsonparametersanddimensions
| fromthe |     |     | form | in thesimpleasynchronousmode |     |     |
| ------- | --- | --- | ---- | ---------------------------- | --- | --- |
ParametersandDimensions
| andfrom | thePDMenu | menuinthe | fullasynchronous |     | mode. |     |
| ------- | --------- | --------- | ---------------- | --- | ----- | --- |
8. ClickDisconnectto disconnectfrom thecurrentCreo Parametricsession.
You canperformthefollowingoperationson parametersfrom theParametersand
Dimensionsform:
• Retrieveallthe parametersof aPARTmodelin thecurrentCreo Parametric
| sessioninsidetheparameter |     |     | table inthe |     |     | form. |
| ------------------------- | --- | --- | ----------- | --- | --- | ----- |
ParametersandDimensions
• Modifythe unit,value,designated status,and description,exceptnameand
typeforeachparameter.
| • Deleteaparameter |             | andallthe          | valuesassociatedwith |           | it. |     |
| ------------------ | ----------- | ------------------ | -------------------- | --------- | --- | --- |
| • Savethe          | updatedlist | ofparametersbackin |                      | themodel. |     |     |
• Savethe listof parametersretrievedfrom themodelin an XMLfile,or read
| theparametersfroma |     | previouslysavedXMLfileinthe |     |     | form. |     |
| ------------------ | --- | --------------------------- | --- | --- | ----- | --- |
You canperformthefollowingoperationson dimensionsfrom theParametersand
Dimensionsform:
• Retrieveallthe dimensionsofa PART modelinthe currentCreoParametric
sessioninsidethedimensions tablein theParametersandDimensions form.
• Modifythe name,nominalvalue, tolerancetype,tolerance value1,and
| tolerancevalue2,exceptID |             |                    | andtypeforeachdimension. |           |     |     |
| ------------------------ | ----------- | ------------------ | ------------------------ | --------- | --- | --- |
| • Savethe                | updatedlist | ofdimensionsbackin |                          | themodel. |     |     |
• Savethe listof dimensionsretrievedfrom themodelin anXML fileor read
| thedimensionsfrom |     | apreviouslysavedXML |     | filein | theform. |     |
| ----------------- | --- | ------------------- | --- | ------ | -------- | --- |
TheParametersandDimensions formcontainingthe parametersretrieved
| froma | PART modelisshownin |     | thefollowingfigure. |     |                                 |     |
| ----- | ------------------- | --- | ------------------- | --- | ------------------------------- | --- |
| 368   |                     |     |                     |     | Creo®ParametricVBAPIUser’sGuide |     |

SampleApplications 369

31
Geometry Traversal
Example1..............................................................................................................371
Example2..............................................................................................................371
Example3..............................................................................................................372
Example4..............................................................................................................372
Example5..............................................................................................................373
Thisappendix illustratesthe relationshipsbetweenfaces, contours, andedges.
ExamplesE-1through E-5showsomesamplepartsand listtheinformation about
theirsurfaces,faces,contours, andedges.
370 Creo®ParametricVBAPIUser’sGuide

| Example                   | 1          |               |     |
| ------------------------- | ---------- | ------------- | --- |
| Thispart has6faces.       |            |               |     |
| • FaceAhas1               | contourand | 4edges.       |     |
| • EdgeE2istheintersection |            | offacesAandB. |     |
| • EdgeE2isacomponentof    |            | contoursC1and | C2. |
| Example                   | 2          |               |     |
| FaceAhas2contoursand6     |            | edges.        |     |
GeometryTraversal 371

| Example | 3   |     |     |     |
| ------- | --- | --- | --- | --- |
Thispart wasextruded froma rectangularcrosssection.Thefeatureon thetop
| wasaddedlater | asan extrudedprotrusion |         | in theshapeofa | semicircle. |
| ------------- | ----------------------- | ------- | -------------- | ----------- |
| • FaceAhas1   | contourand              | 6edges. |                |             |
• FaceBhas2contoursand8edges.
• FaceChas1contourand4edges.
| Example | 4   |     |     |     |
| ------- | --- | --- | --- | --- |
Thispart wasextruded froma crosssectionidenticaltoFaceA. Inthe Sketcher,
thetopboundarywas sketchedwith twolinesandan arc.Thesketch wasthen
| extrudedto  | formthe basepart,as | shown.  |     |     |
| ----------- | ------------------- | ------- | --- | --- |
| • FaceAhas1 | contourand          | 6edges. |     |     |
• FaceBhas1contourand4edges.
• FaceChas1contourand4edges.
| • FaceDhas1 | contourand | 4edges. |                                 |     |
| ----------- | ---------- | ------- | ------------------------------- | --- |
| 372         |            |         | Creo®ParametricVBAPIUser’sGuide |     |

| Example | 5   |     |
| ------- | --- | --- |
Thispart wasextruded froma rectangularcrosssection.Theslot andholefeatures
wereaddedlater.
| • FaceAhas1              | contourand | 8edges. |
| ------------------------ | ---------- | ------- |
| • FaceBhas3contoursand10 |            | edges.  |
GeometryTraversal 373

32
Geometry Representations
SurfaceParameterization.........................................................................................375
EdgeandCurveParameterization............................................................................384
Thisappendix describesthegeometry representationsofthedataused bythe VB
API.
374 Creo®ParametricVBAPIUser’sGuide

| Surface | Parameterization |     |     |
| ------- | ---------------- | --- | --- |
Asurfacein CreoParametriccontains datathatdescribestheboundary ofthe
surface,andapointerto theprimitivesurface onwhichitlies.The primitive
surfaceisa three-dimensionalgeometricsurfaceparameterizedby twovariables
(uandv).The surfaceboundary consistsofclosed loops (contours)ofedges. Each
edgeis attachedto twosurfaces,andeachedgecontainsthe uandvvalues ofthe
portionof theboundarythatitforms forboth surfaces.Surface boundariesare
traversedclockwisearoundthe outsideof asurface,soan edgehasadirectionin
| eachsurface | with respectto | thedirection | oftraversal. |
| ----------- | -------------- | ------------ | ------------ |
Thissectiondescribesthe surfaceparameterization.The surfaces arelistedin
orderof complexity. Forease ofuse, thealphabeticallistingof thedatastructures
isasfollows:
• Coneonpage377
| • Coons             | Patchonpage380        |           |     |
| ------------------- | --------------------- | --------- | --- |
| • Cylinderon        | page376               |           |     |
| • CylindricalSpline | Surface               | onpage383 |     |
| • FilletSurfaceon   | page380               |           |     |
| • GeneralSurface    | ofRevolutiononpage378 |           |     |
| • NURBS             | onpage386             |           |     |
| • Planeon           | page375               |           |     |
| • RuledSurface      | on page379            |           |     |
| • SplineSurface     | on page381            |           |     |
| • TabulatedCylinder | on                    | page379   |     |
| • Toruson           | page377               |           |     |
Plane
Theplaneentityconsists oftwoperpendicularunitvectors (e1ande2), the
| normaltothe | plane(e3),andtheorigin |     | ofthe plane. |
| ----------- | ---------------------- | --- | ------------ |
DataFormat:
GeometryRepresentations 375

| e1[3]     | Unit   | vector, | in the u direction |     |
| --------- | ------ | ------- | ------------------ | --- |
| e2[3]     | Unit   | vector, | in the v direction |     |
| e3[3]     | Normal | to the  | plane              |     |
| origin[3] | Origin | of the  | plane              |     |
Parameterization:
| (x, y, | z) = u * | e1 + v * | e2 + origin |     |
| ------ | -------- | -------- | ----------- | --- |
Cylinder
Thegeneratingcurveofacylinder isa line,parallelto theaxis, ata distanceR
fromtheaxis. Theradialdistance ofapointis constant,andthe heightof thepoint
isv.
DataFormat:
| e1[3]     | Unit   | vector, | in the u direction |     |
| --------- | ------ | ------- | ------------------ | --- |
| e2[3]     | Unit   | vector, | in the v direction |     |
| e3[3]     | Normal | to the  | plane              |     |
| origin[3] | Origin | of the  | cylinder           |     |
| radius    | Radius | of the  | cylinder           |     |
Parameterization:
| (x, y, | z) = radius | * [cos(u) | * e1 + sin(u) | * e2] + |
| ------ | ----------- | --------- | ------------- | ------- |
|        | v * e3      | + origin  |               |         |
EngineeringNotes:
Forthe cylinder,cone, torus,andgeneralsurface ofrevolution,a localcoordinate
systemisused thatconsistsofthree orthogonalunitvectors(e1, e2, ande3)and
anorigin.The curvelies inthe planeofe1 ande3, andisrotated inthe direction
frome1 to e2.The usurfaceparameter determinestheangle ofrotation,and the
vparameterdeterminesthe positionofthe pointonthe generatingcurve.
| 376 |     |     |     | Creo®ParametricVBAPIUser’sGuide |
| --- | --- | --- | --- | ------------------------------- |

Cone
Thegeneratingcurveofaconeis alineatanangle alphato theaxis ofrevolution
thatintersectsthe axisattheorigin. Thevparameter isthe heightofthepoint
alongtheaxis,andthe radialdistance ofthe pointisv *tan(alpha).
DataFormat:
| e1[3]     | Unit vector, |            | in the | u direction |             |
| --------- | ------------ | ---------- | ------ | ----------- | ----------- |
| e2[3]     | Unit vector, |            | in the | v direction |             |
| e3[3]     | Normal       | to the     | plane  |             |             |
| origin[3] | Origin       | of the     | cone   |             |             |
| alpha     | Angle        | between    | the    | axis        | of the cone |
|           | and the      | generating |        | line        |             |
Parameterization:
| (x, y, | z) = v * tan(alpha) |       | *   | [cos(u) | * e1 +   |
| ------ | ------------------- | ----- | --- | ------- | -------- |
|        | sin(u)              | * e2] | + v | * e3    | + origin |
Torus
Thegeneratingcurveofatorus isan arcof radiusR2with itscenteratadistance
R1from theorigin.The startingpointof thegeneratingarcislocated atadistance
R1+ R2from theorigin, inthe directionofthe firstvector ofthelocal coordinate
system.Theradialdistance ofapoint onthetorus isR1 +R2 *cos(v), andthe
|          |                  |     |        | revolutionisR2 | * sin(v). |
| -------- | ---------------- | --- | ------ | -------------- | --------- |
| heightof | thepointalongthe |     | axisof |                |           |
DataFormat:
| e1[3]     | Unit vector, |        | in the | u direction |        |
| --------- | ------------ | ------ | ------ | ----------- | ------ |
| e2[3]     | Unit vector, |        | in the | v direction |        |
| e3[3]     | Normal       | to the | plane  |             |        |
| origin[3] | Origin       | of the | torus  |             |        |
| radius1   | Distance     | from   | the    | center      | of the |
GeometryRepresentations 377

|     | generating |     | arc to | the | axis | of  |
| --- | ---------- | --- | ------ | --- | ---- | --- |
revolution
| radius2 | Radius | of the | generating |     | arc |     |
| ------- | ------ | ------ | ---------- | --- | --- | --- |
Parameterization:
| (x, y, | z) = (R1 + | R2 *  | cos(v)) | *           | [cos(u) | * e1 + |
| ------ | ---------- | ----- | ------- | ----------- | ------- | ------ |
|        | sin(u)     | * e2] | +       | R2 * sin(v) |         | * e3 + |
origin
| General | Surface | of  | Revolution |     |     |     |
| ------- | ------- | --- | ---------- | --- | --- | --- |
Ageneralsurfaceof revolutioniscreated byrotating acurveentity, usuallya
spline,aroundan axis.Thecurveis evaluatedatthe normalizedparameterv,and
theresulting pointisrotatedaround theaxisthrough an angleu.The surfaceof
revolutiondatastructureconsists ofa localcoordinatesystemand acurve
structure.
DataFormat:
| e1[3]     | Unit       | vector, | in          | the u | direction     |     |
| --------- | ---------- | ------- | ----------- | ----- | ------------- | --- |
| e2[3]     | Unit       | vector, | in          | the v | direction     |     |
| e3[3]     | Normal     | to      | the plane   |       |               |     |
| origin[3] | Origin     | of      | the surface |       | of revolution |     |
| curve     | Generating |         | curve       |       |               |     |
Parameterization:
| curve(v) | = (c1, c2, | c3)    | is     | a point      | on  | the curve.                      |
| -------- | ---------- | ------ | ------ | ------------ | --- | ------------------------------- |
| (x, y,   | z) = [c1 * | cos(u) | -      | c2 * sin(u)] |     | * e1 +                          |
|          | [c1 *      | sin(u) | +      | c2 * cos(u)] |     | * e2 +                          |
|          | c3 *       | e3 +   | origin |              |     |                                 |
| 378      |            |        |        |              |     | Creo®ParametricVBAPIUser’sGuide |

| Ruled | Surface |     |     |     |     |
| ----- | ------- | --- | --- | --- | --- |
Aruledsurfaceisthe surfacegeneratedby interpolatinglinearlybetween
correspondingpointsoftwocurveentities. Theu coordinateis thenormalized
parameteratwhichboth curvesareevaluated,andthe vcoordinateisthelinear
parameterbetweenthe twopoints.Thecurvesarenot definedinthe local
coordinatesystem ofthepart, sotheresulting pointmustbe transformedby the
| localcoordinatesystemof |     |     | thesurface. |     |     |
| ----------------------- | --- | --- | ----------- | --- | --- |
DataFormat:
| e1[3]     | Unit   | vector,    | in        | the u direction |     |
| --------- | ------ | ---------- | --------- | --------------- | --- |
| e2[3]     | Unit   | vector,    | in        | the v direction |     |
| e3[3]     | Normal | to         | the plane |                 |     |
| origin[3] | Origin | of         | the ruled | surface         |     |
| curve_1   | First  | generating |           | curve           |     |
| curve_2   | Second | generating |           | curve           |     |
Parameterization:
| (x', y',  | z') is   | the point |         | in local coordinates. |             |
| --------- | -------- | --------- | ------- | --------------------- | ----------- |
| (x', y',  | z') =    | (1 - v)   | * C1(u) | + v *                 | C2(u)       |
| (x, y,    | z) = x'  | * e1 +    | y' *    | e2 + z' *             | e3 + origin |
| Tabulated | Cylinder |           |         |                       |             |
Atabulatedcylinder iscalculatedby projectingacurvelinearlythrough space.
Thecurveisevaluatedattheu parameter,andthe zcoordinateisoffset bythe v
parameter. Theresulting pointis expressedin localcoordinatesandmustbe
transformedby thelocal coordinatesystem tobe expressedin part coordinates.
DataFormat:
| e1[3] | Unit | vector, | in  | the u direction |     |
| ----- | ---- | ------- | --- | --------------- | --- |
GeometryRepresentations 379

| e2[3]     | Unit vector, | in               | the v | direction |     |
| --------- | ------------ | ---------------- | ----- | --------- | --- |
| e3[3]     | Normal       | to the plane     |       |           |     |
| origin[3] | Origin       | of the tabulated |       | cylinder  |     |
| curve     | Generating   | curve            |       |           |     |
Parameterization:
| (x', y', | z') is the | point     | in local | coordinates. |             |
| -------- | ---------- | --------- | -------- | ------------ | ----------- |
| (x', y', | z') = C(u) | + (0,     | 0, v)    |              |             |
| (x, y,   | z) = x' *  | e1 + y' * | e2 +     | z' *         | e3 + origin |
| Coons    | Patch      |           |          |              |             |
ACoonspatch isusedto blendsurfaces together.For example,you wouldusea
Coonspatch atacornerwherethreefillets(eachof adifferentradius) meet.
DataFormat:
| le_curve           |         | u =    | 0 boundary |             |                                 |
| ------------------ | ------- | ------ | ---------- | ----------- | ------------------------------- |
| ri_curve           |         | u =    | 1 boundary |             |                                 |
| dn_curve           |         | v =    | 0 boundary |             |                                 |
| up_curve           |         | v =    | 1 boundary |             |                                 |
| point_matrix[2][2] |         | Corner | points     |             |                                 |
| uvder_matrix[2][2] |         | Corner | mixed      | derivatives |                                 |
| Fillet             | Surface |        |            |             |                                 |
| 380                |         |        |            |             | Creo®ParametricVBAPIUser’sGuide |

Afilletsurface isfound wherearound orafilletis placedona curvededge, oron
anedgewith non-constantarcradii.Onastraight edge,acylinder wouldbe used
| torepresent |     | thefillet. |     |     |     |     |     |     |
| ----------- | --- | ---------- | --- | --- | --- | --- | --- | --- |
DataFormat:
| pnt_spline |     | P(v)   | spline | running | along         |         | the u = | 0 boundary |
| ---------- | --- | ------ | ------ | ------- | ------------- | ------- | ------- | ---------- |
| ctr_spline |     | C(v)   | spline | along   | the           | centers | of      | the        |
|            |     | fillet | arcs   |         |               |         |         |            |
| tan_spline |     | T(v)   | spline | of      | unit tangents |         | to the  |            |
|            |     | axis   | of the | fillet  | arcs          |         |         |            |
Parameterization:
| R(v)    | = P(v) | - C(v) |        |          |        |     |        |     |
| ------- | ------ | ------ | ------ | -------- | ------ | --- | ------ | --- |
| (x,y,z) | =      | C(v)   | + R(v) | * cos(u) | + T(v) | X   | R(v) * |     |
sin(u)
| Spline | Surface |     |     |     |     |     |     |     |
| ------ | ------- | --- | --- | --- | --- | --- | --- | --- |
Theparametricspline surfaceis anonuniformbicubic splinesurface thatpasses
througha gridwith tangentvectors givenateachpoint.Thegrid iscurvilinearin
uvspace.
|     | Use | thisfor | bicubicblending |     | betweencornerpoints. |     |     |     |
| --- | --- | ------- | --------------- | --- | -------------------- | --- | --- | --- |
DataFormat:
| u_par_arr[]    |     |     | Point          | parameters,    |             | in the  | u       |     |
| -------------- | --- | --- | -------------- | -------------- | ----------- | ------- | ------- | --- |
|                |     |     | direction,     | of             | size        | Nu      |         |     |
| v_par_arr[]    |     |     | Point          | parameters,    |             | in the  | v       |     |
|                |     |     | direction,     | of             | size        | Nv      |         |     |
| point_arr[][3] |     |     | Array          | of interpolant |             | points, | of      |     |
|                |     |     | size Nu        | x Nv           |             |         |         |     |
| u_tan_arr[][3] |     |     | Array          | of u tangent   |             | vectors |         |     |
|                |     |     | at interpolant |                | points,     |         | of size |     |
|                |     |     | Nu x Nv        |                |             |         |         |     |
| v_tan_arr[][3] |     |     | Array          | of v tangent   |             | vectors | at      |     |
|                |     |     | interpolant    |                | points,     | of      | size    |     |
|                |     |     | Nu x Nv        |                |             |         |         |     |
| uvder_arr[][3] |     |     | Array          | of mixed       | derivatives |         | at      |     |
|                |     |     | interpolant    |                | points,     | of      | size    |     |
GeometryRepresentations 381

|     |     | Nu x Nv |     |     |
| --- | --- | ------- | --- | --- |
EngineeringNotes:
| • Allowsfora | unique3x3polynomial         |     |     | aroundeverypatch. |
| ------------ | --------------------------- | --- | --- | ----------------- |
| • Thereis    | secondordercontinuityacross |     |     | patchboundaries.  |
• Thepoint andtangentvectors representthe orderingofan arrayof[i][j],
whereuvarieswith I,andvvarieswith j.Inwalkingthrough thepoint_
arr[][3], youwillfind thattheinnermostvariablerepresentingv(j)
variesfirst.
| NURBS | Surface |     |     |     |
| ----- | ------- | --- | --- | --- |
TheNURBS surfaceis definedbybasis functions(inu andv),expandablearrays
| ofknots,weights, | andcontrolpoints. |     |     |     |
| ---------------- | ----------------- | --- | --- | --- |
DataFormat:
| deg[2]      |     | Degree    | of the     | basis     |
| ----------- | --- | --------- | ---------- | --------- |
|             |     | functions | (in        | u and v)  |
| u_par_arr[] |     | Array     | of knots   | on the    |
|             |     | parameter | line       | u         |
| v_par_arr[] |     | Array     | of knots   | on the    |
|             |     | parameter | line       | v         |
| wghts[]     |     | Array     | of weights | for       |
|             |     | rational  | NURBS,     | otherwise |
NULL
| c_point_arr[][3] |     | Array | of control | points |
| ---------------- | --- | ----- | ---------- | ------ |
Definition:
382 Creo®ParametricVBAPIUser’sGuide

| k =    | degree in u      |                |                   |     |     |
| ------ | ---------------- | -------------- | ----------------- | --- | --- |
| l =    | degree in v      |                |                   |     |     |
| N1 =   | (number of knots | in u) -        | (degree in u) - 2 |     |     |
| N2 =   | (number of knots | in v) -        | (degree in v) - 2 |     |     |
| Bi,k = | basis function   | in u           |                   |     |     |
| Bj, l  | = basis function | in v           |                   |     |     |
| wij =  | weights          |                |                   |     |     |
| Ci, j  | = control        | points (x,y,z) | * wi,j            |     |     |
EngineeringNotes:
| Theweights                       | andc_points_arrarraysrepresentmatricesof |                 |                       | size |     |
| -------------------------------- | ---------------------------------------- | --------------- | --------------------- | ---- | --- |
| wghts[N1+1][N2+1]andc_points_arr |                                          |                 | [N1+1][N2+1].Elements |      | of  |
| thematricesarepackedintoarraysin |                                          | row-majororder. |                       |      |     |
| Cylindrical                      | Spline                                   | Surface         |                       |      |     |
Thecylindricalspline surfaceis anonuniformbicubic splinesurface thatpasses
througha gridwith tangentvectors givenateachpoint.Thegrid iscurvilinearin
modelingspace.
DataFormat:
| e1[3] | x' vector | of the local | coordinate |     |     |
| ----- | --------- | ------------ | ---------- | --- | --- |
system
| e2[3] | y' vector | of the local | coordinate |     |     |
| ----- | --------- | ------------ | ---------- | --- | --- |
system
| e3[3]                   | z' vector | of the local      | coordinate  |     |     |
| ----------------------- | --------- | ----------------- | ----------- | --- | --- |
|                         | system,   | which corresponds | to the      |     |     |
|                         | axis of   | revolution of     | the surface |     |     |
| GeometryRepresentations |           |                   |             |     | 383 |

| origin[3] |     | Origin | of the | local |     | coordinate |     |     |
| --------- | --- | ------ | ------ | ----- | --- | ---------- | --- | --- |
system
| splsrf           |     | Spline                | surface     |                 | data        | structure           |              |        |
| ---------------- | --- | --------------------- | ----------- | --------------- | ----------- | ------------------- | ------------ | ------ |
| Thesplinesurface |     | datastructurecontains |             |                 |             | thefollowingfields: |              |        |
| u_par_arr[]      |     |                       | Point       |                 | parameters, |                     | in the       |        |
|                  |     |                       | u           | direction,      |             | of                  | size Nu      |        |
| v_par_arr[]      |     |                       | Point       |                 | parameters, |                     | in the       |        |
|                  |     |                       | v           | direction,      |             | of                  | size Nv      |        |
| point_arr[][3]   |     |                       | Array       |                 | of points,  |                     | in           |        |
|                  |     |                       | cylindrical |                 |             | coordinates,        |              |        |
|                  |     |                       | of          | size            | Nu          | x Nv.               | The array    |        |
|                  |     |                       | components  |                 |             | are                 | as follows:  |        |
|                  |     |                       |             | point_arr[i][0] |             |                     | -            | Radius |
|                  |     |                       |             | point_arr[i][1] |             |                     | -            | Theta  |
|                  |     |                       |             | point_arr[i][2] |             |                     | -            | Z      |
| u_tan_arr[][3]   |     |                       | Array       |                 | of u        | tangent             | vectors.     |        |
|                  |     |                       | in          | cylindrical     |             |                     | coordinates, |        |
|                  |     |                       | of          | size            | Nu          | x Nv                |              |        |
| v_tan_arr[][3]   |     |                       | Array       |                 | of v        | tangent             | vectors,     |        |
|                  |     |                       | in          | cylindrical     |             |                     | coordinates, |        |
|                  |     |                       | of          | size            | Nu          | x Nv                |              |        |
| uvder_arr[][3]   |     |                       | Array       |                 | of mixed    |                     | derivatives, |        |
|                  |     |                       | in          | cylindrical     |             |                     | coordinates, |        |
|                  |     |                       | of          | size            | Nu          | x Nv                |              |        |
EngineeringNotes:
Ifthesurfaceis representedin cylindrical coordinates(r,theta, z),thelocal
| coordinatesystem |     | values(x', |     | y',z')are |     | interpretedasfollows: |     |     |
| ---------------- | --- | ---------- | --- | --------- | --- | --------------------- | --- | --- |
| x' = r           | cos | (theta)    |     |           |     |                       |     |     |
| y' = r           | sin | (theta)    |     |           |     |                       |     |     |
z' = z
Acylindricalsplinesurface canbe obtained,for example,bycreatingasmooth
| rotationalblend |     | (shownin |     | thefigureon |     | thepreviouspage). |     |     |
| --------------- | --- | -------- | --- | ----------- | --- | ----------------- | --- | --- |
Insomecases,youcanreplace acylindrical splinesurfacewith asurface suchasa
plane,cylinder,or cone.For example,inthe figure,thecylindricalspline surface
| S was | replacedwithacone(r1= |     |     |     | r2,r3 |     | =r4,and | r1≠r3). |
| ----- | --------------------- | --- | --- | --- | ----- | --- | ------- | ------- |
1
Ifareplacementcannotbe done(such asfor thesurfaceS in thefigure(ra≠rb
0
| orrc | ≠rd)),leaveitas |       | acylindricalspline |     |                  |     | surfacerepresentation. |     |
| ---- | --------------- | ----- | ------------------ | --- | ---------------- | --- | ---------------------- | --- |
| Edge | and             | Curve |                    |     | Parameterization |     |                        |     |
Thisparameterization representsedges(line,arc,andspline) aswellasthecurves
| (line,arc,spline,and |     |     | NURBS)withinthe |     |     | surfaces. |     |     |
| -------------------- | --- | --- | --------------- | --- | --- | --------- | --- | --- |
Thissectiondescribesedgesand curves,arranged inorder ofcomplexity.For ease
| ofuse,the | alphabeticallistingis |     |     |     | asfollows: |     |     |                                 |
| --------- | --------------------- | --- | --- | --- | ---------- | --- | --- | ------------------------------- |
| 384       |                       |     |     |     |            |     |     | Creo®ParametricVBAPIUser’sGuide |

• Arconpage385
| • Lineon   | page385   |     |     |     |     |
| ---------- | --------- | --- | --- | --- | --- |
| • NURBS    | onpage386 |     |     |     |     |
| • Splineon | page385   |     |     |     |     |
Line
DataFormat:
| end1[3] | Starting | point | of  | the line |     |
| ------- | -------- | ----- | --- | -------- | --- |
| end2[3] | Ending   | point | of  | the line |     |
Parameterization:
| (x, y, | z) = (1 - | t) * | end1 | + t * end2 |     |
| ------ | --------- | ---- | ---- | ---------- | --- |
Arc
Thearcentityis definedbyaplane in whichthearclies.The arciscenteredatthe
origin,andisparameterizedby theangleof rotationfrom thefirst planeunit
vectorin thedirectionof thesecondplane vector.The startand endangle
parametersof thearcandthe radiusarealsogiven.The directionofthe arcis
counterclockwiseifthe startangle islessthantheendangle, otherwiseitis
clockwise.
DataFormat:
| vector1[3]  | First   | vector |           | that defines | the          |
| ----------- | ------- | ------ | --------- | ------------ | ------------ |
|             | plane   | of     | the       | arc          |              |
| vector2[3]  | Second  |        | vector    | that defines | the          |
|             | plane   | of     | the       | arc          |              |
| origin[3]   | Origin  |        | that      | defines the  | plane        |
|             | of      | the    | arc       |              |              |
| start_angle | Angular |        | parameter | of           | the starting |
point
| end_angle | Angular |     | parameter | of  | the ending |
| --------- | ------- | --- | --------- | --- | ---------- |
point
| radius | Radius |     | of the | arc. |     |
| ------ | ------ | --- | ------ | ---- | --- |
Parameterization:
| t' (the | unnormalized     |     | parameter) | is            |     |
| ------- | ---------------- | --- | ---------- | ------------- | --- |
| (1 -    | t) * start_angle |     | +          | t * end_angle |     |
| (x, y,  | z) = radius      | *   | [cos(t')   | * vector1     | +   |
| sin(t') | * vector2]       |     | + origin   |               |     |
Spline
Thesplinecurve entityis anonuniform cubicspline,definedby aseriesofthree-
dimensionalpoints,tangentvectors ateachpoint,andan arrayofunnormalized
splineparametersateachpoint.
GeometryRepresentations 385

DataFormat:
par_arr[] Array of spline parameters
(t) at each point.
pnt_arr[][3] Array of spline interpolant points
tan_arr[][3] Array of tangent vectors at
each point
Parameterization:
x,y,and zareaseriesof uniquecubic functions,onepersegment,fully
determinedbythe startingandending points,andtangentsof eachsegment.
Letp_max bethe parameterof thelast splinepoint.Then, t,theunnormalized
parameter,is t*p_max.
Locatetheithspline segmentsuch that:
par_arr[i] < t' < par_arr[i+1]
(Ift< 0ort >+1,usethefirst orlast segment.)
t0 = (t' - par_arr[i]) / (par_arr[i+1] - par_arr[i])
t1 = (par_arr[i+1] - t') / (par_arr[i+1] - par_arr[i])
NURBS
TheNURBS (nonuniform rationalB-spline) curveisdefinedby expandablearrays
ofknots,weights, andcontrolpoints.
DataFormat:
degree Degree of the basis function
params[] Array of knots
weights[] Array of weights for rational
NURBS, otherwise NULL.
c_pnts[][3] Array of control points
Definition:
386 Creo®ParametricVBAPIUser’sGuide

| k = degree  | of basis  | function   |     |
| ----------- | --------- | ---------- | --- |
| N = (number | of knots) | - (degree) | - 2 |
wi = weights
| Ci = control        | points    | (x, y, z)           | * wi |
| ------------------- | --------- | ------------------- | ---- |
| Bi,k = basis        | functions |                     |      |
| Bythis equation,the | numberof  | controlpointsequals | N+1. |
References:
Faux,I.D.,M.J. Pratt. ComputationalGeometryforDesignandManufacture. Ellis
HarwoodPublishers,1983.
Mortenson,M.E.GeometricModeling.JohnWiley&Sons, 1985.
GeometryRepresentations 387

Index
| A        |                | B            |     |
| -------- | -------------- | ------------ | --- |
| Accuracy |                | B-splines    |     |
|          |                | description, | 221 |
| getting  | andsetting,155 |              |     |
ActionListener
| descriptionofthe | class,21 |     |     |
| ---------------- | -------- | --- | --- |
C
| events,           | 268 |            |     |
| ----------------- | --- | ---------- | --- |
| feature-level,274 |     | Cells      |     |
| session-level,269 |     | accessing, | 264 |
|                   |     | Children,  | 194 |
| solid-level,      | 272 |            |     |
Circles,221
| types, | 268 |     |     |
| ------ | --- | --- | --- |
Clear
UI command,270
window,178
ActionSourceinterface,268
| Activate   |            | Close        |     |
| ---------- | ---------- | ------------ | --- |
| window,178 |            | window,178   |     |
| Add        |            | Colors,47    |     |
| items to   | alayer,191 | Commands     |     |
|            |            | designating, | 81  |
Allocate
|     |     | compatibility | ofdeprecatedmethods, |
| --- | --- | ------------- | -------------------- |
simplifiedrepresentations,317
44
Applications
|                    |     | Composite                | curves |
| ------------------ | --- | ------------------------ | ------ |
| hierarchy,         | 23  |                          |        |
| Arcs               |     | description,             | 221    |
| description,       | 221 | Cones                    |        |
| representation,385 |     | class representation,225 |        |
geometryrepresentation,377
Area
Configurationoptions, 46
| surface, | 227 |     |     |
| -------- | --- | --- | --- |
Contours
Arrows,221
evaluating,224
Assemblies
| coordinatesystems,183 |        | traversing,                | 220                    |
| --------------------- | ------ | -------------------------- | ---------------------- |
| creating,             | 155    | Contours,                  | locatingin amodel, 224 |
| hierarchy,            | 250    | Coons patches              |                        |
| structure             | of,250 | geometryrepresentation,380 |                        |
Coordinatesystems,229
Axes,229
assemblies,183
evaluating,229
|     |     | datum,   | 183                             |
| --- | --- | -------- | ------------------------------- |
|     |     | drawing, | 183                             |
|     |     | Drawing  | View,183                        |
| 388 |     |          | Creo®ParametricVBAPIUser’sGuide |

| evaluating,229                |                   |              | featurepattern,198               |            |                  |
| ----------------------------- | ----------------- | ------------ | -------------------------------- | ---------- | ---------------- |
| screen,182                    |                   |              | models,                          | 94         |                  |
| section,183                   |                   |              | simplifiedrepresentations,317    |            |                  |
| solid, 182                    |                   |              | deprecatedmethods                |            | compatibility,44 |
| window,182                    |                   |              | Depth                            |            |                  |
| Coordinatetransformations,181 |                   |              | selection,61                     |            |                  |
| Copy                          |                   |              | Descriptors                      |            |                  |
| models,                       | 94                |              | model, 88                        |            |                  |
| Create                        |                   |              | Designating                      |            |                  |
| assembly,                     | 155               |              | command,82                       |            |                  |
| buttons,                      | 77                |              | commands,81                      |            |                  |
| family                        | tablecolumns,265  |              | icon, 81                         |            |                  |
| family                        | tableinstance,263 |              | Designatingcommands,81           |            |                  |
| layer,191                     |                   |              | Detail.DetailEntityInstructions  |            |                  |
| local group,198               |                   |              | interface                        |            |                  |
| material,                     | 167               |              | description,                     | 134        |                  |
| menus,77                      |                   |              | Dimension2D.Dimension2Dinterface |            |                  |
| parameters,                   | 234               |              | description,                     | 118        |                  |
| part, 155                     |                   |              | Dimensions,240                   |            |                  |
| simplifiedrepresentations,317 |                   |              | information,240                  |            |                  |
| UDFs,                         | 201               |              | tolerances,241                   |            |                  |
| window,177                    |                   |              | Display                          |            |                  |
| CreateInteractively           |                   | DefinedUDFs, | modelin                          | window,177 |                  |
| 203                           |                   |              | models,                          | 94         |                  |
| CreatingUDFs,                 |                   | 202          | selection,62                     |            |                  |
| CreoParametric                |                   |              | Display status                   |            |                  |
| accessing,                    | 48                |              | of layers,191                    |            |                  |
| Curves,221                    |                   |              | Documentation                    |            |                  |
| datastructures,               |                   | 384          | seeAPIWizard,41                  |            |                  |
| determiningthetype,           |                   | 221          | Drawing                          |            |                  |
| t parameter,221               |                   |              | models                           |            |                  |
| types,                        | 221               |              | codeexample,115                  |            |                  |
| reserved,                     | 221               |              | sheets                           |            |                  |
| Cylinders,225                 |                   |              | example,113                      |            |                  |
| geometryrepresentation,376    |                   |              | transformations,186              |            |                  |
| spline                        | surfaces,383      |              | views                            |            |                  |
| tabulated,                    | 225               |              | codeexample,115                  |            |                  |
geometryrepresentation,379
E
D
Edges,221
| Delete |     |     | determiningthetype, |     | 221 |
| ------ | --- | --- | ------------------- | --- | --- |
Index 389

| evaluating,222            |            |           | accessing,                 | 194             |     |
| ------------------------- | ---------- | --------- | -------------------------- | --------------- | --- |
| t parameter,221           |            |           | creating,                  | 196             |     |
| traversing,               | 220        |           | failed,                    | 194             |     |
| types, 221                |            |           | groups,                    | 194, 198        |     |
| reserved,                 | 221        |           | identifiers,               | 194             |     |
| Ellipses,221              |            |           | information,195            |                 |     |
| Erase                     |            |           | operations,196             |                 |     |
| models,                   | 94         |           | parents,                   | 194             |     |
| Evaluation                |            |           | patterns,198               |                 |     |
| axes, 229                 |            |           | read-only,                 | 195             |     |
| contour,                  | 224        |           | resuming,                  | 196             |     |
| coordinatesystem,229      |            |           | suppressing,               | 196             |     |
| edge, 222                 |            |           | user-defined,201           |                 |     |
| point,229                 |            |           | Files                      |                 |     |
| surface,                  | 226        |           | exporting,                 | 277             |     |
| Examples                  |            |           | message,48                 |                 |     |
| creatingdrawing           |            | views,115 | contents,49                |                 |     |
| listingviews,             | 118        |           | naming                     | restrictions,48 |     |
| normalizingacoordinate    |            |           | plotting,                  | 303             |     |
| transformationmatrix,186  |            |           | Filletsurfaces             |                 |     |
| retrieving                | amodel,    | 91        | geometryrepresentation,380 |                 |     |
| settingangulartolerances, |            | 242       |                            |                 |     |
| using drawing             | sheets,113 |           |                            |                 |     |
G
| visiting | theitems | in asimplified |                |               |     |
| -------- | -------- | -------------- | -------------- | ------------- | --- |
|          |          |                | Generalsurface | ofrevolution, | 378 |
representation,318
| Exceptions      |     |     | Genericmodel |     |     |
| --------------- | --- | --- | ------------ | --- | --- |
| ActionListener, |     | 23  | getting,263  |     |     |
| Export          |     |     | Geometry     |     |     |
| files, 277      |     |     | solidedge,   | 223 |     |
terms,220
|     |     |     | traversal, | 220 |     |
| --- | --- | --- | ---------- | --- | --- |
F
|     |     |     | Groups, 198,201 |     |     |
| --- | --- | --- | --------------- | --- | --- |
Faces
| traversing, | 220 |     |     |     |     |
| ----------- | --- | --- | --- | --- | --- |
H
Familytables,263
| cells,264 |     |     | Hierarchy    |     |     |
| --------- | --- | --- | ------------ | --- | --- |
| columns   |     |     | application, | 23  |     |
Highlight
| accessing, | 264 |     |     |     |     |
| ---------- | --- | --- | --- | --- | --- |
selections,62
instances
| accessing, | 263 |     |     |     |     |
| ---------- | --- | --- | --- | --- | --- |
| symbols,   | 264 |     |     |     |     |
I
Features
Information
| 390 |     |     |     | Creo®ParametricVBAPIUser’sGuide |     |
| --- | --- | --- | --- | ------------------------------- | --- |

| Drawing,                 | 108 |     | Massproperties, | 164        |
| ------------------------ | --- | --- | --------------- | ---------- |
| Initialize               |     |     | Materials,167   |            |
| ActionListeners,         | 22  |     | Matrix          |            |
| Install                  |     |     | codeexample,186 |            |
| See J-LinkInstalling,365 |     |     | Matrix3D        | object,186 |
| Installation             |     |     | Message         | files,48   |
| See J-LinkInstalling,365 |     |     | contents,49     |            |
| Interactiveselection,60  |     |     | restrictions,   | 48         |
| InteractivelyDefinedUDFs |     |     | Message         | window     |
| create,203               |     |     | readingfrom,    | 50         |
|                          |     |     | writingto,      | 50         |
ModelItems
K
duplicating,191
| Keywords   |     |     | evaluating,229  |     |
| ---------- | --- | --- | --------------- | --- |
| instanceof |     |     | getting,189     |     |
| using,     | 221 |     | information,190 |     |
|            |     |     | types,          | 189 |
Models
L
|                |     |     | descriptors,  | 88  |
| -------------- | --- | --- | ------------- | --- |
| Layers,191     |     |     | Drawing       |     |
| operations,191 |     |     | Obtaining,108 |     |
Lines
|              |     |     | exporting, | 277 |
| ------------ | --- | --- | ---------- | --- |
| description, | 221 |     |            |     |
getting,88
representation,385
operations,94
styles,47
retrieving,89
| Lists                 |     |     | codeexample,91                |     |
| --------------------- | --- | --- | ----------------------------- | --- |
| of children,194       |     |     | Modify                        |     |
| of currentwindows,177 |     |     | simplifiedrepresentations,318 |     |
of layeritems,191
of materials,167
N
of ModelItems,189
| of patternmembers,194, |                   | 198 | Normalize          |     |
| ---------------------- | ----------------- | --- | ------------------ | --- |
| of rowsin              | a familytable,263 |     | matrix,186         |     |
| of subitems,189        |                   |     | NURBS              |     |
| of views,180           |                   |     | representation,386 |     |
| of windows,177         |                   |     | surface,382        |     |
Localgroups
| creating, | 198 |     |     |     |
| --------- | --- | --- | --- | --- |
O
Locks,263
Open
file, 89
M
Operations
| Macros,46 |     |     | Drawing, | 109 |
| --------- | --- | --- | -------- | --- |
Index 391

| feature,196 |     |     |     | Principalcurve, | 227 |
| ----------- | --- | --- | --- | --------------- | --- |
layer,191
| model, 94 |     |     |     |     |     |
| --------- | --- | --- | --- | --- | --- |
R
| solid, 155, | 311 |     |     |              |     |
| ----------- | --- | --- | --- | ------------ | --- |
| view, 181   |     |     |     | Refresh      |     |
| window,178, | 312 |     |     | window,178,  | 312 |
| Outlines    |     |     |     | Regenerate   |     |
| contour,    | 224 |     |     | events,269   |     |
|             |     |     |     | solids, 155, | 311 |
Remove
P
|                                      |                 |     |     | items froma                   | layer,191   |
| ------------------------------------ | --------------- | --- | --- | ----------------------------- | ----------- |
| Parameters,234                       |                 |     |     | Rename                        |             |
| information,237                      |                 |     |     | models,                       | 94          |
| ParamValueobjects,233                |                 |     |     | Repaint                       |             |
| Parents,194                          |                 |     |     | events,269                    |             |
| Parts,167                            |                 |     |     | window,178,                   | 312         |
| creating,                            | 155             |     |     | Reset                         |             |
| Patternleaders,194,                  |                 | 198 |     | view, 181                     |             |
| Patterns,198                         |                 |     |     | Restrictions                  |             |
| pfcExceptions.                       |                 |     |     | on textmessagefiles,48        |             |
| XToolkitDrawingCreateErrors          |                 |     |     | Retrieve                      |             |
| description,                         | 107             |     |     | geometryof                    | asimplified |
| pfcModel.Models                      |                 |     |     | representation,316            |             |
| information,91                       |                 |     |     | graphicsof                    | asimplified |
| pfcSelect.Selection.GetSelItemmethod |                 |     |     | representation,316            |             |
| getting aModelItem,                  |                 | 189 |     | material,                     | 167         |
| pfcSelect.SelectionOptionsargument   |                 |     |     | model                         |             |
| table ofstrings,60                   |                 |     |     | codeexample,91                |             |
| Planes,225                           |                 |     |     | simplifiedrepresentations,316 |             |
| geometryrepresentation,375           |                 |     |     | view, 180                     |             |
| Plot,303                             |                 |     |     | Revolvedsurfaces,             | 225         |
| Points,229                           |                 |     |     | Rotate                        |             |
| evaluating,229                       |                 |     |     | view, 181                     |             |
| Polygons,221                         |                 |     |     | Ruled surfaces,225            |             |
| PopupMenu                            |                 |     |     | geometryrepresentation,379    |             |
| Adding tothe                         | GraphicsWindow, |     | 84  |                               |             |
| Using Trailfilestodetermine          |                 |     |     |                               |             |
S
names,84
Save
Popupmenus
| Adding,85    |     |     |     | models,                     | 94                              |
| ------------ | --- | --- | --- | --------------------------- | ------------------------------- |
| PopupMenus,  | 83  |     |     | view, 181                   |                                 |
| Accessing,84 |     |     |     | Screen coordinatesystem,182 |                                 |
|              |     |     |     | Selection,                  | 60                              |
| 392          |     |     |     |                             | Creo®ParametricVBAPIUser’sGuide |

| accessingdata,            | 61  |     | revolved,225            |      |     |
| ------------------------- | --- | --- | ----------------------- | ---- | --- |
| controllingdisplay,       |     | 62  | ruled,                  | 225, | 379 |
| Sheets                    |     |     | spline,381              |      |     |
| Drawing,                  | 110 |     | traversing,             |      | 220 |
| Simplifiedrepresentations |     |     | types,                  | 225  |     |
| adding items,318          |     |     | UV parameterization,225 |      |     |
| creating,                 | 317 |     |                         |      |     |
deleting,317
T
| items, | 318 |     |     |     |     |
| ------ | --- | --- | --- | --- | --- |
t parameter
| extractinginformation |     | from,317 |                |           |     |
| --------------------- | --- | -------- | -------------- | --------- | --- |
| modifying,318         |     |          | description,   |           | 221 |
| retrieving            |     |          | Table          |           |     |
| geometry,316          |     |          | creating,      | 125       |     |
| graphics,316          |     |          | drawing        | cells,125 |     |
| utilities,320         |     |          | retrieving,126 |           |     |
| Solids                |     |          | selecting,125  |           |     |
Table.Tableinterface
| accuracy,                | 155 |     |                            |     |     |
| ------------------------ | --- | --- | -------------------------- | --- | --- |
| coordinatesystem,182     |     |     | description,               |     | 124 |
| geometrytraversal,       |     | 189 | Tabulatedcylinders,225     |     |     |
| getting asolidobject,155 |     |     | geometryrepresentation,379 |     |     |
| information,155          |     |     | Text,221                   |     |     |
| mass properties,164      |     |     | messagefiles,              |     | 48  |
| operations,155,          |     | 311 | Tolerance,                 | 241 |     |
Tolerances
Splines
setting
| cylindrical        | splinesurface,383 |     |                      |                       |     |
| ------------------ | ----------------- | --- | -------------------- | --------------------- | --- |
| description,       | 221               |     | example,242          |                       |     |
| representation,385 |                   |     | Torii,225            |                       |     |
| surface,           | 381               |     | Torus,               | 377                   |     |
| Status             |                   |     | Transformations,181, |                       | 183 |
| feature,195        |                   |     | solidto              | coordinatesystemdatum |     |
| layer,191          |                   |     | coordinates,186      |                       |     |
|                    |                   |     | solidto              | screencoordinates,    | 185 |
Surfaces,225
|                       |            |     | in                 | adrawing, | 186 |
| --------------------- | ---------- | --- | ------------------ | --------- | --- |
| cylindrical           | spline,383 |     |                    |           |     |
| datastructures,       |            | 375 | Traversal          |           |     |
| evaluating,226        |            |     | of asolidblock,220 |           |     |
| area, 227             |            |     | of geometry,220    |           |     |
| evaluatingparameters, |            | 227 |                    |           |     |
fillet
U
geometryrepresentation,380
UDFs, 201
| general surfaceofrevolution, |     | 378 |           |       |     |
| ---------------------------- | --- | --- | --------- | ----- | --- |
| NURBS                        |     |     | creating, | 201   |     |
| geometryrepresentation,382   |     |     | User’s    | Guide |     |
documentation
Index 393

online, 41
Utilities
simplifiedrepresentations,320
V
Values
ParamValue, 233
View2D.View2Dinterface
description, 113
Views
Drawing, 113
getting aview object,180
list of,180
listing
example,118
operations,181
retrieving, 180
saving,181
Visibility,195
Visit
simplifiedrepresentations,317
W
Windowcoordinatesystem,182
Windows, 177
activating,178
clearing, 178
closing, 178
creating, 177
flush, 178
operations,178, 312
repainting, 178
Write
to themessage window,50
394 Creo®ParametricVBAPIUser’sGuide