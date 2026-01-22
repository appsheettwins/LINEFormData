const btn = document.getElementById('btn');
const liffId = '2008940385-theIq9K7'; // ใส่ LIFF ID ที่ระบุ
let userId = '';

function main(){
    // 1. Initial LIFF
    liff.init({ liffId: liffId }).then(() => {
        if (!liff.isLoggedIn()) {
            liff.login();
        } else {
            // 2. ดึง Profile เมื่อ Login แล้ว
            liff.getProfile().then((profile) => {            
                userId = profile.userId;
                console.log("LIFF Profile:", profile);
            }).catch((err) => console.error("Get Profile Error:", err));
        }
    }).catch((err) => console.error("LIFF Init Error:", err));
}

main();

async function send(e){     
    e.preventDefault();
    
    // แสดงสถานะกำลังส่ง
    btn.disabled = true;
    btn.value = "กำลังบันทึก...";

    const firstname = document.getElementById('firstname').value;
    const lastname = document.getElementById('lastname').value;
    const email = document.getElementById('email').value;    
    
    // URL ของ Google Apps Script ที่คุณให้มา
    const scriptUrl = 'https://script.google.com/macros/s/AKfycbzcEIJiQ5eszD5bqqXxJlt8TSuC6-UkwSijsWFikOnJ1trNPF0iOKjpegrYFKxjEZyYeQ/exec'; 

    try {
        // ใช้ axios.post ตามโครงสร้างเดิมของคุณ
        // หมายเหตุ: การส่งไป Google Apps Script บางครั้งอาจติด CORS 
        // หากติดแนะนำให้ใช้ fetch(scriptUrl, {method: 'POST', mode: 'no-cors', ...}) แทน
        const result = await axios.post(scriptUrl, JSON.stringify({
            firstname: firstname,
            lastname: lastname,
            email: email,
            userId: userId
        }));        
        
        alert("บันทึกข้อมูลเรียบร้อยแล้ว!");
        liff.closeWindow(); // ปิดหน้าต่าง LIFF เมื่อสำเร็จ
        
    } catch (error) {
        console.error("Error sending data:", error);
        alert("เกิดข้อผิดพลาดในการส่งข้อมูล");
    } finally {
        btn.disabled = false;
        btn.value = "Submit";
    }
}
