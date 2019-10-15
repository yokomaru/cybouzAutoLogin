'サイボウズ自動ログイン
'------------------------------
const USER_NAME = "XXXXX" 'cybousに登録されているユーザ名()
const PASS_WORD = "XXXXX" 'cybousパスワード
const SPAN_LAST_NAME = "XXX" '自分の上の名前(漢字)
const SPAN_FIRST_NAME = "XXX" '自分の下の名前(漢字)
'------------------------------

const LOGIN_URL = "https://XXXXXXX.s.cybozu.com/login"
const LOGIN_HREF ="https://XXXXXXX.s.cybozu.com/o/"

Set ie = CreateObject("InternetExplorer.Application")
	ie.Visible = true
	ie.Navigate LOGIN_URL
	WaitIE

	'既にログイン済みかどうかを確認する
	If (NameCheck() = true)Then
		WaitIE
		LinkCheckCybozu
	else
		WaitIE
		LoginCybozu
		Wscript.Sleep 1000
		LinkCheckCybozu
	end if
	
	function NameCheck
		 'spanオブジェクト内に自分の名前のオブジェクトがあれば、パスワードは入力済み(パスワード入力後の画面)
		 for each obj in ie.Document.getElementsByTagName("span")
	 		if (obj.innerText = SPAN_LAST_NAME & "　" & SPAN_FIRST_NAME ) Then
				NameCheck = true
				Exit for
			else
				NameCheck = false
	 		end if
	 	next
	End Function
	
	Sub LoginCybozu
		'パスワード入力画面でユーザ名とパスワードを入力してログインボタンをクリック
		ie.Document.all.item("username").Value = USER_NAME
		ie.Document.all.item("password").Value = PASS_WORD
		ie.Document.getElementsByClassname("login-button")(0).Click
	End Sub
	
	Sub LinkCheckCybozu
		WaitIE
		'aオブジェクトのhrefにサイボウズのトップ画面のURLがあればトップ画面に遷移
		for each obj in ie.Document.getElementsByTagName("a")
			if (obj.href = LOGIN_HREF) Then
				ie.Navigate obj.href
				Exit for
			End if
		next
	End Sub
	
	Sub WaitIE
		Do While ie.Busy = True Or ie.readystate <> 4
			WScript.sleep 100
		Loop
			Wscript.Sleep 100
	End Sub