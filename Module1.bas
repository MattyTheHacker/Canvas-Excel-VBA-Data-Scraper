Attribute VB_Name = "Module1"
Sub DownloadCanvasData()

    Dim token As String
    Dim courses_url As String
    Dim quiz_url As String
    Dim course_request As New WinHttpRequest
    Dim quiz_request As New WinHttpRequest
    Dim course_ids As ArrayList
    Set course_ids = New ArrayList
    Dim course_id As Long


    ' hard code token as variable until we have oauth2 working
    token = "obviously-not-my-actual-token-you-numpty"

    ' set course url
    courses_url = "https://canvas.bham.ac.uk/api/v1/courses/"

    ' set up request to get all courses
    course_request.Open "GET", courses_url, False

    ' set authorisation header
    course_request.setRequestHeader "Authorization", "Bearer " & token

    ' send request
    course_request.Send

    ' check if request was successful
    If course_request.Status <> 200 Then
        MsgBox "ERROR: " & course_request.Status & " " & course_request.statusText
        Exit Sub
    End If

    ' get response
    Dim response As Object
    Set response = JSONConverter.ParseJson(course_request.ResponseText)

    ' loop through courses and put "id" and "name" into the demo sheet
    Dim course As Object

    Sheets("courses").Range("A1").Value = "id"
    Sheets("courses").Range("B1").Value = "name"

    Dim i As Integer
    i = 2

    For Each course In response
        course_ids.Add course("id")
        Sheets("courses").Range("A" & i).Value = course("id")
        Sheets("courses").Range("B" & i).Value = course("name")
        i = i + 1
    Next

    ' set quiz url
    quiz_url = "https://canvas.bham.ac.uk/api/v1/users/234863/courses/56073/assignments"

    ' set up request to get all quizzes
    quiz_request.Open "GET", quiz_url, False

    ' set authorisation header
    quiz_request.setRequestHeader "Authorization", "Bearer " & token

    ' send request
    quiz_request.Send

    ' check if request was successful
    If quiz_request.Status <> 200 Then
        MsgBox "ERROR: " & quiz_request.Status & " " & quiz_request.statusText
        Exit Sub
    End If

    ' get response
    Set response = JSONConverter.ParseJson(quiz_request.ResponseText)

    Debug.Print JSONConverter.ConvertToJson(quiz_request.ResponseText)

    ' loop through quizzes and put "id" and "name" into the quizzes sheet
    Dim quiz As Object

    Sheets("quizzes").Range("A1").Value = "id"
    Sheets("quizzes").Range("B1").Value = "name"
    Sheets("quizzes").Range("C1").Value = "points_possible"

    i = 2

    For Each quiz In response
        Sheets("quizzes").Range("A" & i).Value = quiz("id")
        Sheets("quizzes").Range("B" & i).Value = quiz("name")
        Sheets("quizzes").Range("C" & i).Value = quiz("points_possible")
        i = i + 1
    Next

End Sub
