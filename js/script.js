document.addEventListener('DOMContentLoaded',function(){
    const burger=document.getElementById('burger');
    const navMenu=document.querySelector('.nav-menu');
    if(burger&&navMenu){burger.addEventListener('click',function(){navMenu.classList.toggle('open')})}
});
