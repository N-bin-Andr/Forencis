


let maxLenght;

function getPearsonMaxLenght(lenght, man, dry){
/*Определение прижизненного роста индивидуума мужского пола по методике Pearson
(категория низкого роста по условной рубрикации Мартина)/
*/
//man=true; dry=true;
	function => maxLenght =  78.664+2.376*(0.1*lenght-0.96)-2;
//man=true; dry=false;		
	function => maxLenght = 78.807+2.376*(0.1*lenght-0.96)-2;
//man=false; dry=true;
	function => maxLenght = 74.774 + 2.352 * (0.1*lenght-0.87)-1.26;
//man=false; dry=false;	
	function => maxLenght = 75.369+2.352*(0.1*lenght-0.87)-1.26;
	
//	return какоето-число (только число!!)
}

function checkPearsonMaxLenght(lenght, man){
/*определение правильности выбранной методики (по условной рубрикации Мартина)*/	
//man=true;
	if (man==true ? m:w);
	let m = function() =>(lenght<163.9 ? false:true);
//man=false;
	let w = function() =>(lenght<152.9 ? false:true);
}

function PearsonMaxLenght(check)=>{if(check==false ? "методика не подходит": maxLenght)}
/*Вывод результатов определения прижизненного роста индивидуума мужского пола по методике Pearson
(категория низкого роста по условной рубрикации Мартина)/
*/


/*
Public Function Pearson(tmpMaxLenght As Single, Man As Boolean, Optional Dry As Boolean = True) As String
'Определение прижизненного роста индивидуума мужского пола по методике Pearson
'(категория низкого роста по условной рубрикации Мартина)
    Dim tmp As Single
        'Пол мужской или женский:
    If Man = True Then
        'Состояние кости: "сухая": Dry = True; "влажная": Dry = False
        If Dry = True Then
            tmp = 78.664 + 2.376 * (0.1 * tmpMaxLenght - 0.96) - 2 'для перевода в см: 0,1*tmpMaxLenght
        Else: tmp = 78.807 + 2.376 * (0.1 * tmpMaxLenght - 0.96) - 2
        End If
        'определение правильности выбранной методики (по условной рубрикации Мартина)
            If tmp > 163.9 Then
                Pearson = "методика не подходит"
            Else
                Pearson = Format(tmp, "#00.0") & "см."
            End If
    Else:
        If Dry = True Then
            tmp = 74.774 + 2.352 * (0.1 * tmpMaxLenght - 0.87) - 1.26 'для перевода в см: 0,1*tmpMaxLenght
        Else: tmp = 75.369 + 2.352 * (0.1 * tmpMaxLenght - 0.87) - 1.26
        End If
        'определение правильности выбранной методики (по условной рубрикации Мартина)
        If tmp > 152.9 Then
            Pearson = "методика не подходит"
        Else
            Pearson = Format(tmp, "#00.0") & "см."
        End If
    End If
Debug.Print "Pearson = " & Pearson
End Function
*/



