Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Threading
Imports System.Threading.Tasks

Namespace CreoRegenAppVB
    Friend Class OperationResult
        Friend Property Success As Boolean
        Friend Property Message As String

        Friend Shared Function Ok(message As String) As OperationResult
            Return New OperationResult With {.Success = True, .Message = message}
        End Function

        Friend Shared Function Fail(message As String) As OperationResult
            Return New OperationResult With {.Success = False, .Message = message}
        End Function
    End Class

    Friend Class ParameterItem
        Friend Property Name As String
        Friend Property Discr As Integer
        Friend Property ValueText As String
        Friend Property Description As String
    End Class

    Friend Class ParameterQueryResult
        Friend Property Success As Boolean
        Friend Property Message As String
        Friend Property Items As List(Of ParameterItem)

        Friend Shared Function Ok(items As List(Of ParameterItem), message As String) As ParameterQueryResult
            Return New ParameterQueryResult With {.Success = True, .Message = message, .Items = items}
        End Function

        Friend Shared Function Fail(message As String) As ParameterQueryResult
            Return New ParameterQueryResult With {.Success = False, .Message = message, .Items = New List(Of ParameterItem)()}
        End Function
    End Class

    Friend Class CreoConnectionSettings
        Friend Property CreoStartCommand As String = "parametric"
        Friend Property TextDirectory As String = "."
        Friend Property ConnectionUser As String = Nothing
        Friend Property ConnectionDisplay As String = Nothing
        Friend Property ConnectionHost As String = Nothing
        Friend Property ConnectionPort As String = Nothing
    End Class

    Friend Class CreoSessionService
        Implements IDisposable

        Private Shared ReadOnly _extensionKindMap As Dictionary(Of String, CreoModelKind) =
            New Dictionary(Of String, CreoModelKind)(StringComparer.OrdinalIgnoreCase) From {
                {"PRT", CreoModelKind.Part},
                {"ASM", CreoModelKind.Assembly},
                {"DRW", CreoModelKind.Drawing},
                {"MFG", CreoModelKind.Part},
                {"FRM", CreoModelKind.Drawing},
                {"LAY", CreoModelKind.Drawing},
                {"DGM", CreoModelKind.Drawing},
                {"SEC", CreoModelKind.Part},
                {"SKF", CreoModelKind.Part}
            }

        Private Shared ReadOnly _defaultExtensionOrder As String() = {
            "PRT", "ASM", "DRW", "MFG", "FRM", "LAY", "DGM", "SEC", "SKF"
        }

        Private ReadOnly _queue As BlockingCollection(Of Action)
        Private ReadOnly _worker As Thread

        Private _disposed As Boolean
        Private _asyncConnection As Object
        Private _session As Object

        Friend Sub New()
            _queue = New BlockingCollection(Of Action)()
            _worker = New Thread(AddressOf ProcessQueue)
            _worker.Name = "CreoApiWorker"
            _worker.IsBackground = True
            _worker.SetApartmentState(ApartmentState.STA)
            _worker.Start()
        End Sub

        Friend Function EnsureConnectedAsync(settings As CreoConnectionSettings) As Task(Of OperationResult)
            If settings Is Nothing Then
                Throw New ArgumentNullException("settings")
            End If

            Return Enqueue(Function()
                               If _session IsNot Nothing Then
                                   Return OperationResult.Ok("Session already connected.")
                               End If

                               Dim creator = CreatePfcObject("pfcAsyncConnection")

                               Try
                                   _asyncConnection = InvokeMethod(creator, "Connect", settings.ConnectionUser, settings.ConnectionDisplay, settings.ConnectionHost, settings.ConnectionPort)
                                   _session = GetSessionFromConnection(_asyncConnection)
                                   Return OperationResult.Ok("Connected to active Creo session.")
                               Catch
                                   Dim command = BuildStartCommand(settings.CreoStartCommand)
                                   _asyncConnection = InvokeMethod(creator, "Start", command, settings.TextDirectory)
                                   _session = GetSessionFromConnection(_asyncConnection)
                                   Return OperationResult.Ok("Started and connected to new Creo session.")
                               End Try
                           End Function)
        End Function

        Friend Function RetrieveModelAsync(modelInput As String) As Task(Of OperationResult)
            Return Enqueue(Function()
                               If _session Is Nothing Then
                                   Return OperationResult.Fail("No active Creo session. Click Connect first.")
                               End If

                               Dim parsed = ModelNameParser.Parse(modelInput)
                               Dim resolved As ParsedModelName = Nothing
                               Dim model = TryRetrieveModel(parsed, resolved)

                               If model Is Nothing Then
                                   Return OperationResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                               End If

                               InvokeMethod(model, "Display")

                               Dim win = InvokeMethod(_session, "GetModelWindow", model)
                               If win Is Nothing Then
                                   win = InvokeMethod(_session, "CreateModelWindow", model)
                               End If
                               If win IsNot Nothing Then
                                   InvokeMethod(win, "Activate")
                               End If

                               Dim fileName = TryGetString(model, "FileName", "GetFileName")
                               If String.IsNullOrWhiteSpace(fileName) Then
                                   fileName = resolved.FileName
                               End If

                               Return OperationResult.Ok("Retrieved and activated: " & fileName)
                           End Function)
        End Function

        Friend Function RegenerateModelAsync(modelInput As String) As Task(Of OperationResult)
            Return Enqueue(Function()
                               If _session Is Nothing Then
                                   Return OperationResult.Fail("No active Creo session. Click Connect first.")
                               End If

                               Dim parsed = ModelNameParser.Parse(modelInput)
                               Dim resolved As ParsedModelName = Nothing
                               Dim model = TryRetrieveModel(parsed, resolved)
                               If model Is Nothing Then
                                   Return OperationResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                               End If

                               If Not CanRegenerate(model) Then
                                   Dim reportedType = TryGetString(model, "Type", "GetType")
                                   Return OperationResult.Fail("Model type cannot be regenerated (type=" & reportedType & ").")
                               End If

                               Dim modelType As Integer = -1
                               Try
                                   modelType = CInt(GetProperty(model, "Type"))
                               Catch
                               End Try

                               Try
                                   If modelType = 2 OrElse modelType = 4 Then
                                       ' Drawing/Layout regenerate signature takes no arguments.
                                       InvokeMethod(model, "Regenerate")
                                   Else
                                       ' Solid regenerate signature takes IpfcRegenInstructions.
                                       Dim regenFactory = CreatePfcObject("CpfcRegenInstructions")
                                       Dim regenInstructions = InvokeMethod(regenFactory, "Create")
                                       InvokeMethod(model, "Regenerate", regenInstructions)
                                   End If
                               Catch ex As Exception
                                   Try
                                       ActivateModelWindow(model)
                                       RunRegenerationMacros(modelType)
                                   Catch macroEx As Exception
                                       Return OperationResult.Fail("Regeneration failed: " & ex.Message & " (modelType=" & modelType.ToString(CultureInfo.InvariantCulture) & "). Macro fallback failed: " & macroEx.Message)
                                   End Try
                               End Try

                               Dim fileName = TryGetString(model, "FileName", "GetFileName")
                               If String.IsNullOrWhiteSpace(fileName) Then
                                   fileName = resolved.FileName
                               End If

                               Return OperationResult.Ok("Regeneration succeeded: " & fileName)
                           End Function)
        End Function

        Friend Function GetParametersAsync(modelInput As String) As Task(Of ParameterQueryResult)
            Return EnqueueParameters(Function()
                                         If _session Is Nothing Then
                                             Return ParameterQueryResult.Fail("No active Creo session. Click Connect first.")
                                         End If

                                         Dim parsed = ModelNameParser.Parse(modelInput)
                                         Dim resolved As ParsedModelName = Nothing
                                         Dim model = TryRetrieveModel(parsed, resolved)
                                         If model Is Nothing Then
                                             Return ParameterQueryResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                                         End If

                                         Dim paramsObj = GetParametersCollection(model)
                                         If paramsObj Is Nothing Then
                                             Return ParameterQueryResult.Fail("Test Check: Selected model does not expose parameters (paramsObj is null).")
                                         End If

                                         Dim countMsg As String = ""
                                         Dim count = GetCollectionCount(paramsObj, countMsg)

                                         Dim items As New List(Of ParameterItem)()

                                         For i As Integer = 0 To count - 1
                                             Dim param = GetCollectionItem(paramsObj, i)
                                             If param Is Nothing Then
                                                 Continue For
                                             End If

                                             Dim paramValue = GetProperty(param, "Value")
                                             Dim discr = GetParamDiscr(paramValue)

                                             items.Add(New ParameterItem With {
                                                .Name = TryGetString(param, "Name", "GetName"),
                                                .Discr = discr,
                                                .ValueText = GetParamValueText(paramValue, discr),
                                                .Description = TryGetParameterDescription(param)
                                             })
                                         Next

                                         Dim fileName = TryGetString(model, "FileName", "GetFileName")
                                         If String.IsNullOrWhiteSpace(fileName) Then
                                             fileName = resolved.FileName
                                         End If

                                         If items.Count = 0 Then
                                            Return ParameterQueryResult.Fail("Test Check: Found paramsObj. Count debug=[" & countMsg & "] which gave " & count & ". Items loaded: " & items.Count)
                                         End If

                                         Return ParameterQueryResult.Ok(items, "Loaded " & items.Count.ToString(CultureInfo.InvariantCulture) & " parameter(s) from " & fileName)
                                     End Function)
        End Function

        Friend Function SetParameterAsync(modelInput As String, parameterName As String, userInput As String) As Task(Of OperationResult)
            Return Enqueue(Function()
                               If _session Is Nothing Then
                                   Return OperationResult.Fail("No active Creo session. Click Connect first.")
                               End If

                               If String.IsNullOrWhiteSpace(parameterName) Then
                                   Return OperationResult.Fail("Parameter name is empty.")
                               End If

                               Dim parsed = ModelNameParser.Parse(modelInput)
                               Dim resolved As ParsedModelName = Nothing
                               Dim model = TryRetrieveModel(parsed, resolved)
                               If model Is Nothing Then
                                   Return OperationResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                               End If

                               Dim param = FindParameterByName(model, parameterName)
                               If param Is Nothing Then
                                   Return OperationResult.Fail("Parameter not found: " & parameterName)
                               End If

                               Dim currentValue = GetProperty(param, "Value")
                               Dim discr = GetParamDiscr(currentValue)
                               Dim newValue = CreateParamValueFromInput(discr, userInput)
                               If newValue Is Nothing Then
                                   Return OperationResult.Fail("Set failed: Could not create parameter value object for discr=" & discr.ToString(CultureInfo.InvariantCulture) & ". Input: " & userInput)
                               End If

                               Try
                                   SetProperty(param, "Value", newValue)
                               Catch ex As Exception
                                   Return OperationResult.Fail("Error assigning new value to parameter: " & ex.Message)
                               End Try

                               Dim modelType As Integer = -1
                               Try
                                   modelType = CInt(GetProperty(model, "Type"))
                               Catch
                               End Try

                               Try
                                   If modelType = 2 OrElse modelType = 4 Then
                                       InvokeMethod(model, "Regenerate")
                                   Else
                                       Dim regenFactory = CreatePfcObject("CpfcRegenInstructions")
                                       Dim regenInstructions = InvokeMethod(regenFactory, "Create")
                                       InvokeMethod(model, "Regenerate", regenInstructions)
                                   End If
                               Catch
                                   ' Parameter updates still apply even when direct API regen is unavailable.
                                   ActivateModelWindow(model)
                                   RunRegenerationMacros(modelType)
                               End Try

                               Return OperationResult.Ok("Parameter updated: " & parameterName & " = " & userInput)
                           End Function)
        End Function

        Private Function TryRetrieveModel(parsed As ParsedModelName, ByRef resolved As ParsedModelName) As Object
            resolved = Nothing

            For Each candidate In BuildCandidateNames(parsed)
                Try
                    Dim desc = CreateModelDescriptor(candidate)
                    Dim model = InvokeMethod(_session, "RetrieveModel", desc)
                    If model IsNot Nothing Then
                        resolved = candidate
                        Return model
                    End If
                Catch
                End Try
            Next

            Return Nothing
        End Function

        Private Function BuildCandidateNames(parsed As ParsedModelName) As List(Of ParsedModelName)
            Dim candidates As New List(Of ParsedModelName)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim addCandidate =
                Sub(baseName As String, ext As String, kind As CreoModelKind, hasExtension As Boolean)
                    Dim safeBase = baseName.Trim()
                    If safeBase.Length = 0 Then
                        Return
                    End If

                    Dim safeExt = If(ext, String.Empty).Trim().ToUpperInvariant()
                    Dim fileName = If(safeExt.Length = 0, safeBase, safeBase & "." & safeExt.ToLowerInvariant())
                    If seen.Contains(fileName) Then
                        Return
                    End If

                    seen.Add(fileName)
                    candidates.Add(New ParsedModelName With {
                        .Name = safeBase,
                        .Ext = safeExt,
                        .Kind = kind,
                        .FileName = fileName,
                        .HasExplicitExtension = hasExtension
                    })
                End Sub

            If parsed.HasExplicitExtension Then
                addCandidate(parsed.Name, parsed.Ext, parsed.Kind, True)
            Else
                For Each ext In _defaultExtensionOrder
                    addCandidate(parsed.Name, ext, GetKindFromExtension(ext), True)
                Next
            End If

            addCandidate(parsed.Name, String.Empty, CreoModelKind.Unknown, False)
            Return candidates
        End Function

        Private Function GetKindFromExtension(extension As String) As CreoModelKind
            If String.IsNullOrWhiteSpace(extension) Then
                Return CreoModelKind.Unknown
            End If

            Dim ext = extension.Trim().ToUpperInvariant()
            If _extensionKindMap.ContainsKey(ext) Then
                Return _extensionKindMap(ext)
            End If

            Return CreoModelKind.Unknown
        End Function

        Friend Function DisconnectAsync() As Task(Of OperationResult)
            Return Enqueue(Function()
                               If _asyncConnection Is Nothing Then
                                   _session = Nothing
                                   Return OperationResult.Ok("No active connection to close.")
                               End If

                               Try
                                   InvokeMethod(_asyncConnection, "Disconnect")
                               Catch
                               End Try

                               _session = Nothing
                               _asyncConnection = Nothing
                               Return OperationResult.Ok("Disconnected from Creo session.")
                           End Function)
        End Function

        Private Function BuildStartCommand(creoStartCommand As String) As String
            Dim baseCmd = If(String.IsNullOrWhiteSpace(creoStartCommand), "parametric", creoStartCommand.Trim())
            Return baseCmd
        End Function

        Private Function CreateModelDescriptor(parsed As ParsedModelName) As Object
            Dim descCreator = CreatePfcObject("pfcModelDescriptor")

            Try
                Return InvokeMethod(descCreator, "CreateFromFileName", parsed.FileName)
            Catch
            End Try

            If parsed.Kind = CreoModelKind.Unknown Then
                Throw New InvalidOperationException("Unsupported or unknown model type for descriptor: " & parsed.FileName)
            End If

            Return InvokeMethod(descCreator, "Create", GetModelTypeValue(parsed.Kind), parsed.Name, DBNull.Value)
        End Function

        Private Function GetModelTypeValue(kind As CreoModelKind) As Integer
            Select Case kind
                Case CreoModelKind.Assembly
                    Return 0
                Case CreoModelKind.Drawing
                    Return 2
                Case Else
                    Return 1
            End Select
        End Function

        Private Function CanRegenerate(model As Object) As Boolean
            ' Parts, Assemblies, Drawings can all be regenerated.
            Return True
        End Function

        Private Sub ActivateModelWindow(model As Object)
            InvokeMethod(model, "Display")

            Dim win = InvokeMethod(_session, "GetModelWindow", model)
            If win Is Nothing Then
                win = InvokeMethod(_session, "CreateModelWindow", model)
            End If

            If win IsNot Nothing Then
                InvokeMethod(win, "Activate")
            End If
        End Sub

        Private Sub RunRegenerationMacros(modelType As Integer)
            InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdNamedViewsGalSelect` `Default`;")
            InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdRegenAuto`; #CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;")

            If modelType = 2 OrElse modelType = 4 Then
                InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdDwgRegenModel`;#AUTOMATIC;")
                InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdDwgRegenDraft`;")
            Else
                InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdRegenPart`;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;#CURRENT VALS;")
            End If

            InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdViewRefit`;")
            InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdViewRepaint`;")
            InvokeMethod(_session, "RunMacro", "mapkey(continued) ~ Command `ProCmdViewRegenCur`;")
        End Sub

        Private Function GetSessionFromConnection(connection As Object) As Object
            Dim sessionObj = GetProperty(connection, "Session")
            If sessionObj IsNot Nothing Then
                Return sessionObj
            End If

            Return InvokeMethod(connection, "GetSession")
        End Function

        Private Function Enqueue(work As Func(Of OperationResult)) As Task(Of OperationResult)
            If _disposed Then
                Return Task.FromResult(OperationResult.Fail("Service is already disposed."))
            End If

            Dim tcs As New TaskCompletionSource(Of OperationResult)(TaskCreationOptions.RunContinuationsAsynchronously)
            _queue.Add(Sub()
                           Try
                               tcs.TrySetResult(work())
                           Catch ex As Exception
                               tcs.TrySetResult(OperationResult.Fail("CREO ERROR: " & ex.Message))
                           End Try
                       End Sub)
            Return tcs.Task
        End Function

        Private Function EnqueueParameters(work As Func(Of ParameterQueryResult)) As Task(Of ParameterQueryResult)
            If _disposed Then
                Return Task.FromResult(ParameterQueryResult.Fail("Service is already disposed."))
            End If

            Dim tcs As New TaskCompletionSource(Of ParameterQueryResult)(TaskCreationOptions.RunContinuationsAsynchronously)
            _queue.Add(Sub()
                           Try
                               tcs.TrySetResult(work())
                           Catch ex As Exception
                               tcs.TrySetResult(ParameterQueryResult.Fail("CREO ERROR: " & ex.Message))
                           End Try
                       End Sub)
            Return tcs.Task
        End Function

        Private Sub ProcessQueue()
            For Each action In _queue.GetConsumingEnumerable()
                action.Invoke()
            Next
        End Sub

        Private Function CreatePfcObject(className As String) As Object
            Dim progIds = {
                "pfcls." & className,
                "pfcls.C" & className,
                "pfc." & className,
                className
            }

            For Each progId In progIds
                Try
                    Return Activator.CreateInstance(Type.GetTypeFromProgID(progId, throwOnError:=True))
                Catch
                End Try
            Next

            Throw New InvalidOperationException("Unable to create Creo COM object for " & className & ". Run vb_api_register.bat and verify COM registration.")
        End Function

        Private Function GetProperty(instance As Object, propertyName As String) As Object
            Try
                Return CallByName(instance, propertyName, CallType.Get)
            Catch
                Return Nothing
            End Try
        End Function

        Private Sub SetProperty(instance As Object, propertyName As String, value As Object)
            Try
                CallByName(instance, propertyName, CallType.Set, value)
            Catch exSet As Exception
                Try
                    CallByName(instance, propertyName, CallType.Let, value)
                Catch exLet As Exception
                    Throw New Exception("SetProperty failed on " & propertyName & ". Set error: " & exSet.Message & ". Let error: " & exLet.Message)
                End Try
            End Try
        End Sub

        Private Function InvokeMethod(instance As Object, methodName As String, ParamArray args As Object()) As Object
            Select Case args.Length
                Case 0
                    Return CallByName(instance, methodName, CallType.Method)
                Case 1
                    Return CallByName(instance, methodName, CallType.Method, args(0))
                Case 2
                    Return CallByName(instance, methodName, CallType.Method, args(0), args(1))
                Case 3
                    Return CallByName(instance, methodName, CallType.Method, args(0), args(1), args(2))
                Case 4
                    Return CallByName(instance, methodName, CallType.Method, args(0), args(1), args(2), args(3))
                Case Else
                    Throw New NotSupportedException("InvokeMethod currently supports up to 4 arguments for late-bound COM calls.")
            End Select
        End Function

        Private Function TryGetString(instance As Object, propertyName As String, methodName As String) As String
            Dim value As Object = Nothing

            Try
                value = CallByName(instance, propertyName, CallType.Get)
            Catch
            End Try

            If value Is Nothing Then
                Try
                    value = CallByName(instance, methodName, CallType.Method)
                Catch
                End Try
            End If

            If value Is Nothing Then
                Return String.Empty
            End If

            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Function

        Private Function FindParameterByName(model As Object, parameterName As String) As Object
            Try
                Dim param = InvokeMethod(model, "GetParam", parameterName)
                If param IsNot Nothing Then
                    Return param
                End If
            Catch
            End Try

            Dim paramsObj = GetParametersCollection(model)
            If paramsObj Is Nothing Then
                Return Nothing
            End If

            Dim countMsg As String = ""
            Dim count = GetCollectionCount(paramsObj, countMsg)

            For i As Integer = 0 To count - 1
                Dim param = GetCollectionItem(paramsObj, i)
                Dim name = TryGetString(param, "Name", "GetName")
                If String.Equals(name, parameterName, StringComparison.OrdinalIgnoreCase) Then
                    Return param
                End If
            Next

            Return Nothing
        End Function

        Private Function GetParamDiscr(paramValue As Object) As Integer
            If paramValue Is Nothing Then
                Return -1
            End If

            Try
                Dim val = GetProperty(paramValue, "discr")
                If val IsNot Nothing Then Return CInt(val)
            Catch
            End Try

            Try
                Dim val = GetProperty(paramValue, "Discr")
                If val IsNot Nothing Then Return CInt(val)
            Catch
            End Try

            Return -1
        End Function

        Private Function GetParametersCollection(model As Object) As Object
            Try
                Dim paramsObj = InvokeMethod(model, "ListParams")
                If paramsObj IsNot Nothing Then
                    Return paramsObj
                End If
            Catch
            End Try

            Try
                Dim paramsObj = InvokeMethod(model, "ListParameters")
                If paramsObj IsNot Nothing Then
                    Return paramsObj
                End If
            Catch
            End Try

            Try
                Dim paramsObj = GetProperty(model, "Params")
                If paramsObj IsNot Nothing Then
                    Return paramsObj
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function GetCollectionCount(collection As Object, ByRef outMsg As String) As Integer
            outMsg = ""
            If collection Is Nothing Then
                outMsg = "Collection is null."
                Return 0
            End If

            Try
                Dim countVal = GetProperty(collection, "Count")
                If countVal IsNot Nothing Then
                    outMsg = "GetProperty Count succeeded."
                    Return CInt(countVal)
                End If
            Catch ex As Exception
                outMsg = "GetProperty Count exception: " & ex.Message
            End Try

            Try
                Dim countVal = InvokeMethod(collection, "Count")
                If countVal IsNot Nothing Then
                    outMsg = "InvokeMethod Count succeeded."
                    Return CInt(countVal)
                End If
            Catch ex As Exception
                outMsg &= " InvokeMethod Count exception: " & ex.Message
            End Try

            outMsg &= " Could not get Count property."
            Return 0
        End Function

        Private Function GetCollectionItem(collection As Object, zeroBasedIndex As Integer) As Object
            If collection Is Nothing Then
                Return Nothing
            End If

            ' Try default indexed property (like VBA Params(i))
            Try
                Return collection(zeroBasedIndex)
            Catch
            End Try

            ' In VBA, Params(i) is a default property indexer, usually "Item" with CallType.Get
            Try
                Return CallByName(collection, "Item", CallType.Get, zeroBasedIndex)
            Catch
            End Try

            Try
                Return CallByName(collection, "Item", CallType.Method, zeroBasedIndex)
            Catch
            End Try

            Return Nothing
        End Function

        Private Function TryGetParameterDescription(param As Object) As String
            Dim text = TryGetString(param, "Description", "GetDescription")
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            text = TryGetString(param, "Prompt", "GetPrompt")
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            Return String.Empty
        End Function

        Private Function GetParamValueText(paramValue As Object, discr As Integer) As String
            If paramValue Is Nothing Then
                Return String.Empty
            End If

            Try
                Select Case discr
                    Case 0
                        Return Convert.ToString(GetProperty(paramValue, "StringValue"), CultureInfo.InvariantCulture)
                    Case 1
                        Return Convert.ToString(GetProperty(paramValue, "IntValue"), CultureInfo.InvariantCulture)
                    Case 2
                        Return Convert.ToString(GetProperty(paramValue, "BoolValue"), CultureInfo.InvariantCulture)
                    Case 3
                        Return Convert.ToString(GetProperty(paramValue, "DoubleValue"), CultureInfo.InvariantCulture)
                    Case 4
                        Return Convert.ToString(GetProperty(paramValue, "NoteId"), CultureInfo.InvariantCulture)
                    Case Else
                        Return Convert.ToString(paramValue, CultureInfo.InvariantCulture)
                End Select
            Catch
                Return String.Empty
            End Try
        End Function

        Private Function CreateParamValueFromInput(discr As Integer, userInput As String) As Object
            Dim modelItemFactory As Object = Nothing

            Try
                ' We try explicit CMpfcModelItem since that maps to PTC's VB API exact module
                modelItemFactory = Activator.CreateInstance(Type.GetTypeFromProgID("pfcls.CMpfcModelItem", True))
            Catch
            End Try

            If modelItemFactory Is Nothing Then
                Try
                    modelItemFactory = CreatePfcObject("MpfcModelItem")
                Catch
                    Try
                        modelItemFactory = CreatePfcObject("pfcModelItem")
                    Catch
                        Return Nothing
                    End Try
                End Try
            End If

            Try
                Select Case discr
                    Case 0
                        Return InvokeMethod(modelItemFactory, "CreateStringParamValue", userInput)
                    Case 1
                        Dim parsedInt As Integer
                        If Not Integer.TryParse(userInput, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedInt) Then
                            Integer.TryParse(userInput, NumberStyles.Integer, CultureInfo.CurrentCulture, parsedInt)
                        End If
                        Return InvokeMethod(modelItemFactory, "CreateIntParamValue", parsedInt)
                    Case 2
                        Dim parsedBool As Boolean
                        If Not Boolean.TryParse(userInput, parsedBool) Then
                            Dim normalized = userInput.Trim().ToUpperInvariant()
                            parsedBool = (normalized = "1" OrElse normalized = "YES" OrElse normalized = "Y" OrElse normalized = "TRUE")
                        End If
                        Return InvokeMethod(modelItemFactory, "CreateBoolParamValue", parsedBool)
                    Case 3
                        Dim parsedDouble As Double
                        If Not Double.TryParse(userInput, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, parsedDouble) Then
                            Double.TryParse(userInput, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, parsedDouble)
                        End If
                        Return InvokeMethod(modelItemFactory, "CreateDoubleParamValue", parsedDouble)
                    Case Else
                        Return Nothing
                End Select
            Catch
                Return Nothing
            End Try
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then
                Return
            End If

            _disposed = True

            Try
                If _asyncConnection IsNot Nothing Then
                    Try
                        InvokeMethod(_asyncConnection, "Disconnect")
                    Catch
                    End Try
                End If
            Finally
                _session = Nothing
                _asyncConnection = Nothing
            End Try

            _queue.CompleteAdding()
            If _worker.IsAlive Then
                _worker.Join(TimeSpan.FromSeconds(2))
            End If
            _queue.Dispose()
        End Sub
    End Class
End Namespace
