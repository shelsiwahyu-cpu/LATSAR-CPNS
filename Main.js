const email = document.getElementById('email')
const password = document.getElementById('password')
const submit = document.getElementById('submit')

let myemail = "shelsi.wahyu@gmail.com"
let mypassword = "shelshi123"

submit.addEventListener('click',function(){
    let youremail = email.value
    let yourpassword = password.value
    if(youremail == myemail){
    if(mypassword == yourpassword){
          alert ('login succesfully! Mengalihkan ke halaman utama')
            window.location.href = "login.html";
        }else{
            alert('maaf kata sandi anda salah')
        } 
    }else{
        alert('maaf email tidak dikenali')
    }
})

