Office.onReady(() => {
    window.addEventListener("DOMContentLoaded", () => {
        const btn = document.getElementById("createTaskButton");
        if (btn) {
            btn.addEventListener("click", () => {
                console.log("Send button clicked");
                alert("Кнопка работает!");
            });
        }
    });
});
