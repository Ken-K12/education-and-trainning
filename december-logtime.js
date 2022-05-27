/**
 * 12月版
 * Salesforce上の勤怠情報から必要な工数と綱目を取得し、CSVへ保存するプログラム
 * 
 */

//遅延をかます関数
function delay(n){
    return new Promise(function(resolve){
        setTimeout(resolve,n*1000);
    });
}

// メイン処理
async function myAsyncFunction(){
    let str = "";

    //1日から31日までのログ情報を取得
    for (let i=1; i<=31; i++){
        // 9日より前の場合(1桁)
        if(i <= 9){    
            //空白じゃない場合(休日対策) 
            if(document.getElementById(`dailyWorkButton2022-12-0${i}`) != null){
                // 勤務詳細を押下
                document.getElementById(`dailyWorkButton2022-12-0${i}`).click();
                //0.5秒の遅延を発生させる
                await delay(0.5);
                // 日付のHTML情報を取得する
                let date = document.getElementById("empWorkDate").innerHTML;
                // 日付の情報部分を切り出す
                day = date.substring(0,9);
                // 週の情報を切り出す
                let dayOfWeek = date.substring(9, 10);
                // ログ出力
                console.log(`${day}(${dayOfWeek}) : ` + document.getElementById('empTimeLabel3').innerHTML);
                
                //工数登録情報がある場合、strに情報を格納
                if(document.getElementById('empTimeLabel3').innerHTML != ""){
                    str += `${day}(${dayOfWeek}),`
                    str += document.getElementById('empTimeLabel3').innerHTML + '\n';                
                }
                await delay(0.5);
                document.getElementById('empWorkClose').click();
                await delay(0.5);
            }
        }

        // 10日以上の場合(2桁) , 処理は上記と全く同じ
        if(i >= 10){
            if(document.getElementById(`dailyWorkButton2022-12-${i}`) != null){
                document.getElementById(`dailyWorkButton2022-12-${i}`).click();
                await delay(0.5);
                date = document.getElementById("empWorkDate").innerHTML;
                day = date.substring(0,9);
                dayOfWeek = date.substring(10,11);             
                
                console.log(`${day}(${dayOfWeek}) : ` + document.getElementById('empTimeLabel3').innerHTML);
                if(document.getElementById('empTimeLabel3').innerHTML != ""){
                    str += `${day}(${dayOfWeek}),`
                    str += document.getElementById('empTimeLabel3').innerHTML + '\n';                
                }
                await delay(0.5);
                document.getElementById('empWorkClose').click();
                await delay(0.5);
            }
        }     
    }

    //配列に上記の文字列(str)を設定
    let blob =new Blob([str],{type:"text/csv"});
    // リンクを作成 
    let link =document.createElement('a');
    // blobをURL変換してaタグ化
    link.href = URL.createObjectURL(blob);
    // ダウンロードファイル名
    link.download ="December.csv";
    // リンクを押下
    link.click();
}

myAsyncFunction();