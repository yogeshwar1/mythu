Class Student
	private m_name
	public Property Get Name
			Name=m_name
	End Property	
	public Property Let Name(sname)
	m_name=sname
	End Property
	public sub register
		MsgBox "student" &m_name& "Registered"
			End Sub
End Class
Dim s1,s2
Set S1=New student
s1.name="Narender"
Set S2=new Student
s2.name="Vedant"
S1.register
S2.register
