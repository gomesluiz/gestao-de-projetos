 Sub PERT()
Dim tarefa As Task
For Each tarefa In ActiveProject.Tasks
If Not (tarefa Is Nothing) Then
       tarefa.Duration = (tarefa.Duration1 + 4 * tarefa.Duration3 + tarefa.Duration2) / 6
       Else
            MsgBox Prompt:="Erro no cálculo da duração!"
       End If
Next tarefa
End Sub