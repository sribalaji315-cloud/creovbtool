Imports System
Imports System.Collections.Concurrent
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
                               Dim desc = CreateModelDescriptor(parsed)
                               Dim model = InvokeMethod(_session, "RetrieveModel", desc)

                               If model Is Nothing Then
                                   Return OperationResult.Fail("Model not found: " & parsed.Name)
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
                                   fileName = parsed.Name
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
                               Dim desc = CreateModelDescriptor(parsed)
                               Dim model = InvokeMethod(_session, "RetrieveModel", desc)
                               If model Is Nothing Then
                                   Return OperationResult.Fail("Model not found: " & parsed.Name)
                               End If

                               If Not CanRegenerate(model) Then
                                   Dim reportedType = TryGetString(model, "Type", "GetType")
                                   Return OperationResult.Fail("Model is not a solid and cannot be regenerated (type=" & reportedType & ").")
                               End If

                               Try
                                   InvokeMethod(model, "Regenerate", Nothing)
                               Catch
                                   Dim regenInstructions = CreatePfcObject("pfcRegenInstructions")
                                   Dim inst = InvokeMethod(regenInstructions, "Create")
                                   InvokeMethod(model, "Regenerate", inst)
                               End Try

                               Dim fileName = TryGetString(model, "FileName", "GetFileName")
                               If String.IsNullOrWhiteSpace(fileName) Then
                                   fileName = parsed.Name
                               End If

                               Return OperationResult.Ok("Regeneration succeeded: " & fileName)
                           End Function)
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
            Dim modelType = TryGetString(model, "Type", "GetType")
            If String.IsNullOrWhiteSpace(modelType) Then
                Return True
            End If

            Dim upper = modelType.ToUpperInvariant()
            If upper.Contains("DRAW") Then
                Return False
            End If

            Return True
        End Function

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

        Private Function InvokeMethod(instance As Object, methodName As String, ParamArray args As Object()) As Object
            Return CallByName(instance, methodName, CallType.Method, args)
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
