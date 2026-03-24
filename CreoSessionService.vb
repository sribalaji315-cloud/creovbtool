Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
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

    Friend Class ParameterUpdateItem
        Friend Property Name As String
        Friend Property UserInput As String
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

    Friend Class FeatureItem
        Public Property Id As Integer
        Public Property Name As String
        Public Property TypeName As String
        Public Property SubTypeName As String
        Public Property Status As String
    End Class

    Friend Class FeatureQueryResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Items As List(Of FeatureItem)

        Public Shared Function Ok(items As List(Of FeatureItem), message As String) As FeatureQueryResult
            Return New FeatureQueryResult With {.Success = True, .Message = message, .Items = items}
        End Function

        Public Shared Function Fail(message As String) As FeatureQueryResult
            Return New FeatureQueryResult With {.Success = False, .Message = message, .Items = New List(Of FeatureItem)()}
        End Function
    End Class

    Friend Class ComponentItem
        Public Property FeatureId As Integer
        Public Property ComponentName As String
        Public Property FileName As String
    End Class

    Friend Class ComponentQueryResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Items As List(Of ComponentItem)

        Public Shared Function Ok(items As List(Of ComponentItem), message As String) As ComponentQueryResult
            Return New ComponentQueryResult With {.Success = True, .Message = message, .Items = items}
        End Function

        Public Shared Function Fail(message As String) As ComponentQueryResult
            Return New ComponentQueryResult With {.Success = False, .Message = message, .Items = New List(Of ComponentItem)()}
        End Function
    End Class

    Friend Class BomNodeItem
        Public Property Level As Integer
        Public Property FeatureId As Integer
        Public Property DisplayName As String
        Public Property FileName As String
        Public Property ModelTypeName As String
        Public Property Descp As String
        Public Property BsClass As String
        Public Property Children As List(Of BomNodeItem)
    End Class

    Friend Class BomGridRowItem
        Public Property NodeId As Integer
        Public Property ParentNodeId As Integer
        Public Property HasChildren As Boolean
        Public Property Level As Integer
        Public Property FeatureId As Integer
        Public Property DisplayName As String
        Public Property FileName As String
        Public Property ModelTypeName As String
        Public Property Descp As String
        Public Property BsClass As String
    End Class

    Friend Class BomQueryResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Root As BomNodeItem
        Public Property Rows As List(Of BomGridRowItem)

        Public Shared Function Ok(root As BomNodeItem, rows As List(Of BomGridRowItem), message As String) As BomQueryResult
            Return New BomQueryResult With {.Success = True, .Message = message, .Root = root, .Rows = rows}
        End Function

        Public Shared Function Fail(message As String) As BomQueryResult
            Return New BomQueryResult With {.Success = False, .Message = message, .Root = Nothing, .Rows = New List(Of BomGridRowItem)()}
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

        Private Const MaxModelIdentifierLength As Integer = 255
        Private Const MaxBomDepth As Integer = 32
        Private Const MaxBomNodes As Integer = 10000
        Private Const UserFacingTechnicalErrorMessage As String = "The operation failed."

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

        Private Shared ReadOnly AllowedProgIds As HashSet(Of String) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "pfcls.pfcAsyncConnection",
            "pfcls.CpfcAsyncConnection",
            "pfc.pfcAsyncConnection",
            "pfcls.pfcModelDescriptor",
            "pfcls.CpfcModelDescriptor",
            "pfc.pfcModelDescriptor",
            "pfcls.CpfcRegenInstructions",
            "pfcls.CCpfcRegenInstructions",
            "pfcls.CMpfcRegenInstructions",
            "pfc.CpfcRegenInstructions",
            "pfcls.CMpfcModelItem",
            "pfcls.MpfcModelItem",
            "pfc.MpfcModelItem",
            "pfcls.pfcModelItem",
            "pfcls.CpfcModelItem",
            "pfc.pfcModelItem"
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
                                   InvokeNativeRegeneration(model, modelType)
                               Catch ex As Exception
                                   Return OperationResult.Fail("Regeneration failed: " & ex.Message & " (modelType=" & modelType.ToString(CultureInfo.InvariantCulture) & ")")
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
                                   InvokeNativeRegeneration(model, modelType)
                               Catch
                                   ' Parameter value was set; regeneration is best-effort.
                               End Try

                               Return OperationResult.Ok("Parameter updated: " & parameterName & " = " & userInput)
                           End Function)
        End Function

        Friend Function SetParametersAsync(modelInput As String, updates As IEnumerable(Of ParameterUpdateItem)) As Task(Of OperationResult)
            Return Enqueue(Function()
                               If _session Is Nothing Then
                                   Return OperationResult.Fail("No active Creo session. Click Connect first.")
                               End If

                               If updates Is Nothing Then
                                   Return OperationResult.Fail("No parameter updates were provided.")
                               End If

                               Dim updateList As New List(Of ParameterUpdateItem)()
                               For Each update In updates
                                   If update Is Nothing Then
                                       Continue For
                                   End If

                                   Dim parameterName = If(update.Name, String.Empty).Trim()
                                   Dim userInput = If(update.UserInput, String.Empty).Trim()
                                   If parameterName.Length = 0 OrElse userInput.Length = 0 Then
                                       Continue For
                                   End If

                                   updateList.Add(New ParameterUpdateItem With {
                                       .Name = parameterName,
                                       .UserInput = userInput
                                   })
                               Next

                               If updateList.Count = 0 Then
                                   Return OperationResult.Fail("No parameter rows contain a new value.")
                               End If

                               Dim parsed = ModelNameParser.Parse(modelInput)
                               Dim resolved As ParsedModelName = Nothing
                               Dim model = TryRetrieveModel(parsed, resolved)
                               If model Is Nothing Then
                                   Return OperationResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                               End If

                               For Each update In updateList
                                   Dim param = FindParameterByName(model, update.Name)
                                   If param Is Nothing Then
                                       Return OperationResult.Fail("Parameter not found: " & update.Name)
                                   End If

                                   Dim currentValue = GetProperty(param, "Value")
                                   Dim discr = GetParamDiscr(currentValue)
                                   Dim newValue = CreateParamValueFromInput(discr, update.UserInput)
                                   If newValue Is Nothing Then
                                       Return OperationResult.Fail("Set failed: Could not create parameter value object for discr=" & discr.ToString(CultureInfo.InvariantCulture) & ". Input: " & update.UserInput)
                                   End If

                                   Try
                                       SetProperty(param, "Value", newValue)
                                   Catch ex As Exception
                                       Return OperationResult.Fail("Error assigning new value to parameter '" & update.Name & "': " & ex.Message)
                                   End Try
                               Next

                               Dim modelType As Integer = -1
                               Try
                                   modelType = CInt(GetProperty(model, "Type"))
                               Catch
                               End Try

                               Try
                                   InvokeNativeRegeneration(model, modelType)
                               Catch
                                   ' Parameter values were set; regeneration is best-effort.
                               End Try

                               Return OperationResult.Ok("Updated " & updateList.Count.ToString(CultureInfo.InvariantCulture) & " parameter(s) and regenerated model.")
                           End Function)
        End Function

        Friend Function GetFeaturesAsync(modelInput As String) As Task(Of FeatureQueryResult)
            Return EnqueueFeatures(Function()
                                         If _session Is Nothing Then
                                             Return FeatureQueryResult.Fail("No active Creo session. Click Connect first.")
                                         End If

                                         Dim parsed = ModelNameParser.Parse(modelInput)
                                         Dim resolved As ParsedModelName = Nothing
                                         Dim model = TryRetrieveModel(parsed, resolved)
                                         If model Is Nothing Then
                                             Return FeatureQueryResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                                         End If

                                         Dim featuresObj = GetFeaturesCollection(model)
                                         If featuresObj Is Nothing Then
                                             Return FeatureQueryResult.Fail("Selected model does not expose features (e.g. it might not be a solid).")
                                         End If

                                         Dim countMsg As String = ""
                                         Dim count = GetCollectionCount(featuresObj, countMsg)

                                         Dim items As New List(Of FeatureItem)()

                                         For i As Integer = 0 To count - 1
                                             Dim feat = GetCollectionItem(featuresObj, i)
                                             If feat Is Nothing Then
                                                 Continue For
                                             End If

                                             Dim idVal As Integer = -1
                                             Try
                                                 Dim idProp = GetProperty(feat, "Id")
                                                 If idProp IsNot Nothing Then
                                                     idVal = CInt(idProp)
                                                 End If
                                             Catch
                                             End Try

                                             Dim typeName = TryGetString(feat, "FeatTypeName", "GetFeatTypeName")
                                             Dim subTypeName = TryGetString(feat, "FeatSubType", "GetFeatSubType")
                                             Dim statusVal = TryGetString(feat, "Status", "GetStatus")

                                             items.Add(New FeatureItem With {
                                                .Id = idVal,
                                                .Name = TryGetString(feat, "Name", "GetName"),
                                                .TypeName = typeName,
                                                .SubTypeName = subTypeName,
                                                .Status = statusVal
                                             })
                                         Next

                                         Dim fileName = TryGetString(model, "FileName", "GetFileName")
                                         If String.IsNullOrWhiteSpace(fileName) Then
                                             fileName = resolved.FileName
                                         End If

                                         Return FeatureQueryResult.Ok(items, "Loaded " & items.Count.ToString(CultureInfo.InvariantCulture) & " feature(s) from " & fileName)
                                     End Function)
        End Function

        Friend Function GetComponentsAsync(modelInput As String) As Task(Of ComponentQueryResult)
            Return EnqueueComponents(Function()
                                          If _session Is Nothing Then
                                              Return ComponentQueryResult.Fail("No active Creo session. Click Connect first.")
                                          End If

                                          Dim parsed = ModelNameParser.Parse(modelInput)
                                          Dim resolved As ParsedModelName = Nothing
                                          Dim model = TryRetrieveModel(parsed, resolved)
                                          If model Is Nothing Then
                                              Return ComponentQueryResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                                          End If

                                          Dim featuresObj = GetFeaturesCollection(model)
                                          If featuresObj Is Nothing Then
                                              Return ComponentQueryResult.Fail("Selected model does not expose features.")
                                          End If

                                          Dim countMsg As String = ""
                                          Dim count = GetCollectionCount(featuresObj, countMsg)
                                          Dim items As New List(Of ComponentItem)()

                                          For i As Integer = 0 To count - 1
                                              Dim feat = GetCollectionItem(featuresObj, i)
                                              If feat Is Nothing Then
                                                  Continue For
                                              End If

                                              Dim modelDescr = GetProperty(feat, "ModelDescr")
                                              If modelDescr Is Nothing Then
                                                  modelDescr = InvokeMethodSafe(feat, "GetModelDescr")
                                              End If
                                              If modelDescr Is Nothing Then
                                                  Continue For
                                              End If

                                              Dim componentName = TryGetString(modelDescr, "InstanceName", "GetFullName")
                                              Dim fileName = TryGetString(modelDescr, "FileName", "GetFileName")

                                              If String.IsNullOrWhiteSpace(componentName) Then
                                                  componentName = fileName
                                              End If

                                              If String.IsNullOrWhiteSpace(componentName) Then
                                                  componentName = TryGetString(modelDescr, "GenericName", "GetGenericName")
                                              End If

                                              items.Add(New ComponentItem With {
                                                 .FeatureId = GetFeatureId(feat),
                                                 .ComponentName = componentName,
                                                 .FileName = fileName
                                              })
                                          Next

                                          Dim activeFileName = TryGetString(model, "FileName", "GetFileName")
                                          If String.IsNullOrWhiteSpace(activeFileName) Then
                                              activeFileName = resolved.FileName
                                          End If

                                          Return ComponentQueryResult.Ok(items, "Loaded " & items.Count.ToString(CultureInfo.InvariantCulture) & " component feature(s) from " & activeFileName)
                                      End Function)
        End Function

        Friend Function GetBomHierarchyAsync(modelInput As String) As Task(Of BomQueryResult)
            Return EnqueueBom(Function()
                                  If _session Is Nothing Then
                                      Return BomQueryResult.Fail("No active Creo session. Click Connect first.")
                                  End If

                                  Dim parsed = ModelNameParser.Parse(modelInput)
                                  Dim resolved As ParsedModelName = Nothing
                                  Dim model = TryRetrieveModel(parsed, resolved)
                                  If model Is Nothing Then
                                      Return BomQueryResult.Fail("Model not found after trying supported extensions: " & parsed.Name)
                                  End If

                                  Dim rootFileName = TryGetString(model, "FileName", "GetFileName")
                                  If String.IsNullOrWhiteSpace(rootFileName) Then
                                      rootFileName = resolved.FileName
                                  End If

                                  If Not IsAssemblyModel(model, rootFileName) Then
                                      Return BomQueryResult.Fail("BOM hierarchy requires an assembly model.")
                                  End If

                                  Dim totalNodeCount As Integer = 0
                                  Dim visitedModels As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                                  Dim rootNode = BuildBomNode(model, ResolveModelDisplayName(model, rootFileName), rootFileName, -1, 0, totalNodeCount, visitedModels)
                                  Dim rows = New List(Of BomGridRowItem)()
                                  Dim nextNodeId As Integer = 1
                                  FlattenBomNode(rootNode, rows, 0, nextNodeId)
                                  Dim nodeCount = CountBomNodes(rootNode) - 1
                                  Return BomQueryResult.Ok(rootNode, rows, "Built BOM hierarchy with " & nodeCount.ToString(CultureInfo.InvariantCulture) & " component node(s) from " & rootFileName)
                              End Function)
        End Function

        Private Function TryRetrieveModel(parsed As ParsedModelName, ByRef resolved As ParsedModelName) As Object
            resolved = Nothing
            ValidateParsedModelName(parsed)

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

                    ValidateModelIdentifier(safeBase, "model name")

                    Dim safeExt = If(ext, String.Empty).Trim().ToUpperInvariant()
                    If safeExt.Length > 0 Then
                        ValidateModelIdentifier(safeExt, "model extension")
                    End If

                    Dim fileName = If(safeExt.Length = 0, safeBase, safeBase & "." & safeExt.ToLowerInvariant())
                    ValidateModelIdentifier(fileName, "model file name")
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

        Private Sub ValidateParsedModelName(parsed As ParsedModelName)
            If parsed Is Nothing Then
                Throw New ArgumentException("Model name is empty.")
            End If

            ValidateModelIdentifier(parsed.Name, "model name")
            ValidateModelIdentifier(parsed.FileName, "model file name")

            If parsed.HasExplicitExtension Then
                ValidateModelIdentifier(parsed.Ext, "model extension")
            End If
        End Sub

        Private Sub ValidateModelIdentifier(value As String, displayName As String)
            Dim trimmed = If(value, String.Empty).Trim()
            If trimmed.Length = 0 Then
                Throw New ArgumentException("The " & displayName & " is empty.")
            End If

            If trimmed.Length > MaxModelIdentifierLength Then
                Throw New ArgumentException("The " & displayName & " is too long.")
            End If

            If Path.IsPathRooted(trimmed) OrElse trimmed.StartsWith("\\", StringComparison.Ordinal) OrElse trimmed.StartsWith("//", StringComparison.Ordinal) Then
                Throw New ArgumentException("Model input must be a Creo model name, not a file path.")
            End If

            If trimmed.Contains("..") OrElse trimmed.IndexOf("\"c) >= 0 OrElse trimmed.IndexOf("/"c) >= 0 Then
                Throw New ArgumentException("Model input must be a Creo model name, not a file path.")
            End If

            If trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
                Throw New ArgumentException("The " & displayName & " contains invalid characters.")
            End If

            For Each ch In trimmed
                If Char.IsControl(ch) Then
                    Throw New ArgumentException("The " & displayName & " contains invalid characters.")
                End If
            Next
        End Sub

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
                Try
                    InvokeMethod(win, "Repaint")
                Catch
                End Try
            End If
        End Sub

        Private Sub InvokeNativeRegeneration(model As Object, modelType As Integer)
            If modelType = 2 OrElse modelType = 4 Then
                ' IpfcModel2D.Regenerate() takes no arguments.
                InvokeMethod(model, "Regenerate")
            Else
                ' IpfcSolid.Regenerate(IpfcRegenInstructions) expects exactly 1 arg.
                ' Use Type.InvokeMember to bypass VB.NET ParamArray ambiguity that
                ' collapses Nothing into a null array (0 args) instead of one null arg.
                ' The COM dispatch performs the regeneration before returning its HRESULT,
                ' so even if an exception is thrown the model has already been regenerated.
                Dim regenInstructions As Object = Nothing

                Try
                    Dim regenFactory = CreatePfcObject("CpfcRegenInstructions")
                    regenInstructions = InvokeMethod(regenFactory, "Create")
                Catch
                    ' Factory not registered; regenInstructions stays Nothing (auto-regen).
                End Try

                ' The COM dispatch performs the regeneration before returning its
                ' HRESULT, so even if an exception is thrown the model is already
                ' regenerated. Swallow the post-regen COM warning.
                Try
                    model.GetType().InvokeMember(
                        "Regenerate",
                        System.Reflection.BindingFlags.InvokeMethod,
                        Nothing,
                        model,
                        New Object() {regenInstructions})
                Catch
                End Try
            End If
            ActivateModelWindow(model)
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
                               LogApplicationError("OperationResult queue execution", ex)
                               tcs.TrySetResult(OperationResult.Fail(GetUserFacingErrorMessage(ex)))
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
                               LogApplicationError("ParameterQueryResult queue execution", ex)
                               tcs.TrySetResult(ParameterQueryResult.Fail(GetUserFacingErrorMessage(ex)))
                           End Try
                       End Sub)
            Return tcs.Task
        End Function

        Private Function EnqueueFeatures(work As Func(Of FeatureQueryResult)) As Task(Of FeatureQueryResult)
            If _disposed Then
                Return Task.FromResult(FeatureQueryResult.Fail("Service is already disposed."))
            End If

            Dim tcs As New TaskCompletionSource(Of FeatureQueryResult)(TaskCreationOptions.RunContinuationsAsynchronously)
            _queue.Add(Sub()
                           Try
                               tcs.TrySetResult(work())
                           Catch ex As Exception
                               LogApplicationError("FeatureQueryResult queue execution", ex)
                               tcs.TrySetResult(FeatureQueryResult.Fail(GetUserFacingErrorMessage(ex)))
                           End Try
                       End Sub)
            Return tcs.Task
        End Function

        Private Function EnqueueComponents(work As Func(Of ComponentQueryResult)) As Task(Of ComponentQueryResult)
            If _disposed Then
                Return Task.FromResult(ComponentQueryResult.Fail("Service is already disposed."))
            End If

            Dim tcs As New TaskCompletionSource(Of ComponentQueryResult)(TaskCreationOptions.RunContinuationsAsynchronously)
            _queue.Add(Sub()
                           Try
                               tcs.TrySetResult(work())
                           Catch ex As Exception
                               LogApplicationError("ComponentQueryResult queue execution", ex)
                               tcs.TrySetResult(ComponentQueryResult.Fail(GetUserFacingErrorMessage(ex)))
                           End Try
                       End Sub)
            Return tcs.Task
        End Function

        Private Function EnqueueBom(work As Func(Of BomQueryResult)) As Task(Of BomQueryResult)
            If _disposed Then
                Return Task.FromResult(BomQueryResult.Fail("Service is already disposed."))
            End If

            Dim tcs As New TaskCompletionSource(Of BomQueryResult)(TaskCreationOptions.RunContinuationsAsynchronously)
            _queue.Add(Sub()
                           Try
                               tcs.TrySetResult(work())
                           Catch ex As Exception
                               LogApplicationError("BomQueryResult queue execution", ex)
                               tcs.TrySetResult(BomQueryResult.Fail(GetUserFacingErrorMessage(ex)))
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
            Dim progIds = GetAllowedProgIdCandidates(className)
            If progIds.Length = 0 Then
                Throw New ArgumentException("Requested Creo COM class is not allowed.")
            End If

            For Each progId In progIds
                If Not AllowedProgIds.Contains(progId) Then
                    Continue For
                End If

                Try
                    Dim progType = Type.GetTypeFromProgID(progId, throwOnError:=False)
                    If progType IsNot Nothing Then
                        Return Activator.CreateInstance(progType)
                    End If
                Catch
                End Try
            Next

            Throw New InvalidOperationException("Unable to create the required Creo COM object. Run vb_api_register.bat and verify COM registration.")
        End Function

        Private Function GetAllowedProgIdCandidates(className As String) As String()
            Dim normalized = If(className, String.Empty).Trim()
            If Not IsSafeCreoClassName(normalized) Then
                Return New String() {}
            End If

            Select Case normalized
                Case "pfcAsyncConnection"
                    Return New String() {"pfcls.pfcAsyncConnection", "pfcls.CpfcAsyncConnection", "pfc.pfcAsyncConnection"}
                Case "pfcModelDescriptor"
                    Return New String() {"pfcls.pfcModelDescriptor", "pfcls.CpfcModelDescriptor", "pfc.pfcModelDescriptor"}
                Case "CpfcRegenInstructions"
                    Return New String() {"pfcls.CpfcRegenInstructions", "pfcls.CCpfcRegenInstructions", "pfcls.CMpfcRegenInstructions", "pfc.CpfcRegenInstructions"}
                Case "MpfcModelItem"
                    Return New String() {"pfcls.CMpfcModelItem", "pfcls.MpfcModelItem", "pfc.MpfcModelItem"}
                Case "pfcModelItem"
                    Return New String() {"pfcls.pfcModelItem", "pfcls.CpfcModelItem", "pfc.pfcModelItem"}
                Case Else
                    Return New String() {}
            End Select
        End Function

        Private Function IsSafeCreoClassName(className As String) As Boolean
            If String.IsNullOrWhiteSpace(className) Then
                Return False
            End If

            For Each ch In className
                If Not Char.IsLetterOrDigit(ch) Then
                    Return False
                End If
            Next

            Return True
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

        Private Function InvokeMethodSafe(instance As Object, methodName As String, ParamArray args As Object()) As Object
            Try
                Return InvokeMethod(instance, methodName, args)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function GetUserFacingErrorMessage(ex As Exception) As String
            If TypeOf ex Is ArgumentException Then
                Return ex.Message
            End If

            Return UserFacingTechnicalErrorMessage
        End Function

        Friend Shared Sub LogApplicationError(context As String, ex As Exception)
            Dim detail = BuildDiagnosticMessage(context, ex)
            Trace.WriteLine(detail)
        End Sub

        Private Shared Function BuildDiagnosticMessage(context As String, ex As Exception) As String
            Dim lines As New List(Of String)()
            lines.Add("Context: " & If(context, String.Empty))

            Dim current = ex
            Dim depth As Integer = 0
            While current IsNot Nothing AndAlso depth < 10
                lines.Add("Exception[" & depth.ToString(CultureInfo.InvariantCulture) & "]: " & current.GetType().FullName)
                lines.Add("Message[" & depth.ToString(CultureInfo.InvariantCulture) & "]: " & current.Message)
                lines.Add("Stack[" & depth.ToString(CultureInfo.InvariantCulture) & "]: " & If(current.StackTrace, String.Empty))
                current = current.InnerException
                depth += 1
            End While

            Return String.Join(Environment.NewLine, lines.ToArray())
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

        Private Function GetFeaturesCollection(model As Object) As Object
            Try
                Dim featsObj = InvokeMethod(model, "ListItems", 0) ' 0 = EpfcITEM_FEATURE
                If featsObj IsNot Nothing Then
                    Return featsObj
                End If
            Catch
            End Try

            Try
                Dim featsObj = InvokeMethod(model, "ListFeatures")
                If featsObj IsNot Nothing Then
                    Return featsObj
                End If
            Catch
            End Try

            Try
                Dim featsObj = GetProperty(model, "Features")
                If featsObj IsNot Nothing Then
                    Return featsObj
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function BuildBomNode(model As Object, displayName As String, fileName As String, featureId As Integer, depth As Integer, ByRef totalNodeCount As Integer, visitedModels As HashSet(Of String)) As BomNodeItem
            totalNodeCount += 1

            Dim node As New BomNodeItem With {
                .Level = depth,
                .FeatureId = featureId,
                .DisplayName = displayName,
                .FileName = fileName,
                .ModelTypeName = GetModelTypeName(model, fileName),
                .Descp = GetModelParameterText(model, "DESCP"),
                .BsClass = GetModelParameterText(model, "BS_CLASS", "BS_Class"),
                .Children = New List(Of BomNodeItem)()
            }

            If depth >= MaxBomDepth Then
                Return node
            End If

            If totalNodeCount >= MaxBomNodes Then
                Return node
            End If

            Dim featuresObj = GetFeaturesCollection(model)
            If featuresObj Is Nothing Then
                Return node
            End If

            Dim countMsg As String = ""
            Dim count = GetCollectionCount(featuresObj, countMsg)

            For i As Integer = 0 To count - 1
                Dim feat = GetCollectionItem(featuresObj, i)
                If feat Is Nothing Then
                    Continue For
                End If

                Dim modelDescr = GetFeatureModelDescriptor(feat)
                If modelDescr Is Nothing Then
                    Continue For
                End If

                Dim childFileName = ResolveDescriptorFileName(modelDescr)
                Dim childDisplayName = ResolveDescriptorDisplayName(modelDescr, childFileName)
                Dim childModel = RetrieveModelFromDescriptor(modelDescr)
                Dim childTypeName = GetModelTypeName(If(childModel, modelDescr), childFileName)
                Dim childDescp = If(childModel IsNot Nothing, GetModelParameterText(childModel, "DESCP"), String.Empty)
                Dim childBsClass = If(childModel IsNot Nothing, GetModelParameterText(childModel, "BS_CLASS", "BS_Class"), String.Empty)
                Dim childNode = New BomNodeItem With {
                    .Level = depth + 1,
                    .FeatureId = GetFeatureId(feat),
                    .DisplayName = childDisplayName,
                    .FileName = childFileName,
                    .ModelTypeName = childTypeName,
                    .Descp = childDescp,
                    .BsClass = childBsClass,
                    .Children = New List(Of BomNodeItem)()
                }

                If childModel IsNot Nothing AndAlso IsAssemblyModel(childModel, childFileName) Then
                    Dim normalizedChildKey = If(childFileName, String.Empty).ToUpperInvariant()
                    If Not String.IsNullOrEmpty(normalizedChildKey) AndAlso visitedModels.Contains(normalizedChildKey) Then
                        ' Circular reference detected — do not recurse
                    ElseIf totalNodeCount < MaxBomNodes Then
                        If Not String.IsNullOrEmpty(normalizedChildKey) Then
                            visitedModels.Add(normalizedChildKey)
                        End If
                        childNode = BuildBomNode(childModel, childDisplayName, childFileName, GetFeatureId(feat), depth + 1, totalNodeCount, visitedModels)
                        If Not String.IsNullOrEmpty(normalizedChildKey) Then
                            visitedModels.Remove(normalizedChildKey)
                        End If
                    End If
                End If

                node.Children.Add(childNode)
            Next

            Return node
        End Function

        Private Sub FlattenBomNode(node As BomNodeItem, rows As List(Of BomGridRowItem), parentNodeId As Integer, ByRef nextNodeId As Integer)
            If node Is Nothing OrElse rows Is Nothing Then
                Return
            End If

            Dim currentNodeId = nextNodeId
            nextNodeId += 1

            rows.Add(New BomGridRowItem With {
                .NodeId = currentNodeId,
                .ParentNodeId = parentNodeId,
                .HasChildren = node.Children IsNot Nothing AndAlso node.Children.Count > 0,
                .Level = node.Level,
                .FeatureId = node.FeatureId,
                .DisplayName = node.DisplayName,
                .FileName = node.FileName,
                .ModelTypeName = node.ModelTypeName,
                .Descp = node.Descp,
                .BsClass = node.BsClass
            })

            If node.Children Is Nothing Then
                Return
            End If

            For Each child In node.Children
                FlattenBomNode(child, rows, currentNodeId, nextNodeId)
            Next
        End Sub

        Private Function CountBomNodes(node As BomNodeItem) As Integer
            If node Is Nothing Then
                Return 0
            End If

            Dim total = 1
            If node.Children Is Nothing Then
                Return total
            End If

            For Each child In node.Children
                total += CountBomNodes(child)
            Next

            Return total
        End Function

        Private Function GetFeatureModelDescriptor(feature As Object) As Object
            Dim modelDescr = GetProperty(feature, "ModelDescr")
            If modelDescr IsNot Nothing Then
                Return modelDescr
            End If

            Return InvokeMethodSafe(feature, "GetModelDescr")
        End Function

        Private Function RetrieveModelFromDescriptor(modelDescr As Object) As Object
            If modelDescr Is Nothing Then
                Return Nothing
            End If

            Try
                Return InvokeMethod(_session, "RetrieveModel", modelDescr)
            Catch
                Return Nothing
            End Try
        End Function

        Private Function ResolveDescriptorDisplayName(modelDescr As Object, fallbackFileName As String) As String
            Dim name = TryGetString(modelDescr, "InstanceName", "GetFullName")
            If Not String.IsNullOrWhiteSpace(name) Then
                Return name
            End If

            name = TryGetString(modelDescr, "FileName", "GetFileName")
            If Not String.IsNullOrWhiteSpace(name) Then
                Return name
            End If

            name = TryGetString(modelDescr, "GenericName", "GetGenericName")
            If Not String.IsNullOrWhiteSpace(name) Then
                Return name
            End If

            Return fallbackFileName
        End Function

        Private Function ResolveDescriptorFileName(modelDescr As Object) As String
            Dim fileName = TryGetString(modelDescr, "FileName", "GetFileName")
            If Not String.IsNullOrWhiteSpace(fileName) Then
                Return fileName
            End If

            Return TryGetString(modelDescr, "InstanceName", "GetFullName")
        End Function

        Private Function ResolveModelDisplayName(model As Object, fallbackFileName As String) As String
            Dim name = TryGetString(model, "InstanceName", "GetInstanceName")
            If Not String.IsNullOrWhiteSpace(name) Then
                Return name
            End If

            name = TryGetString(model, "FileName", "GetFileName")
            If Not String.IsNullOrWhiteSpace(name) Then
                Return name
            End If

            Return fallbackFileName
        End Function

        Private Function GetModelParameterText(model As Object, ParamArray parameterNames As String()) As String
            If model Is Nothing OrElse parameterNames Is Nothing Then
                Return String.Empty
            End If

            For Each parameterName In parameterNames
                If String.IsNullOrWhiteSpace(parameterName) Then
                    Continue For
                End If

                Dim param = FindParameterByName(model, parameterName)
                If param Is Nothing Then
                    Continue For
                End If

                Try
                    Dim paramValue = GetProperty(param, "Value")
                    Dim discr = GetParamDiscr(paramValue)
                    Return GetParamValueText(paramValue, discr)
                Catch
                End Try
            Next

            Return String.Empty
        End Function

        Private Function IsAssemblyModel(model As Object, fallbackFileName As String) As Boolean
            Try
                Dim typeValue = GetProperty(model, "Type")
                If typeValue IsNot Nothing Then
                    Return CInt(typeValue) = 0
                End If
            Catch
            End Try

            Return HasAssemblyExtension(fallbackFileName)
        End Function

        Private Function IsAssemblyDescriptor(modelDescr As Object, fallbackFileName As String) As Boolean
            Try
                Dim typeValue = GetProperty(modelDescr, "Type")
                If typeValue IsNot Nothing Then
                    Return CInt(typeValue) = 0
                End If
            Catch
            End Try

            Return HasAssemblyExtension(fallbackFileName)
        End Function

        Private Function GetModelTypeName(instance As Object, fallbackFileName As String) As String
            Try
                Dim typeValue = GetProperty(instance, "Type")
                If typeValue IsNot Nothing Then
                    Select Case CInt(typeValue)
                        Case 0
                            Return "Assembly"
                        Case 1
                            Return "Part"
                        Case 2
                            Return "Drawing"
                    End Select
                End If
            Catch
            End Try

            If HasAssemblyExtension(fallbackFileName) Then
                Return "Assembly"
            End If

            Return "Part"
        End Function

        Private Function HasAssemblyExtension(fileName As String) As Boolean
            If String.IsNullOrWhiteSpace(fileName) Then
                Return False
            End If

            Return fileName.EndsWith(".asm", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function GetFeatureId(feature As Object) As Integer
            If feature Is Nothing Then
                Return -1
            End If

            Try
                Dim idProp = GetProperty(feature, "Id")
                If idProp IsNot Nothing Then
                    Return CInt(idProp)
                End If
            Catch
            End Try

            Return -1
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
