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

async function send(e) {
    e.preventDefault();
    
    const btn = document.getElementById('btn');
    btn.disabled = true;
    btn.value = "Saving...";

    const firstname = document.getElementById('firstname').value;
    const lastname = document.getElementById('lastname').value;
    const email = document.getElementById('email').value;

    const scriptUrl = 'https://script.google.com/macros/s/AKfycbzcEIJiQ5eszD5bqqXxJlt8TSuC6-UkwSijsWFikOnJ1trNPF0iOKjpegrYFKxjEZyYeQ/exec';

    const payload = {
        firstname: firstname,
        lastname: lastname,
        email: email,
        userId: userId // ตัวแปร userId ที่ได้จาก liff.getProfile()
    };

    try {
        // ใช้ fetch แทน axios เพื่อลดปัญหา CORS กับ Google Apps Script
        await fetch(scriptUrl, {
            method: 'POST',
            mode: 'no-cors', // สำคัญมากสำหรับการส่งไป GAS
            cache: 'no-cache',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        // เนื่องจากโหมด no-cors เราจะไม่เห็น response status 200 
        // ให้ถือว่าถ้าไม่เข้า catch คือส่งข้อมูลออกไปแล้ว
        alert("ส่งข้อมูลเรียบร้อยแล้ว!");
        liff.closeWindow(); //

    } catch (error) {
        console.error("Error:", error);
        alert("เกิดข้อผิดพลาด: " + error.message);
    } finally {
        btn.disabled = false;
        btn.value = "Submit";
    }
}
