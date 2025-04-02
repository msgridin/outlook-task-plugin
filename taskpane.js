Office.onReady(() => {
    document.getElementById("sendButton").addEventListener("click", createTaskRequest);
});

function createTaskRequest() {
    Office.context.mailbox.item.body.getAsync("text", function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const bodyText = result.value;

            fetch("https://httpbin.org/post", { // заменишь на свой API
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    subject: Office.context.mailbox.item.subject,
                    body: bodyText
                })
            })
                .then(res => res.json())
                .then(data => {
                    console.log("Server response:", data);
                    alert("Success !");
                })
                .catch(error => {
                    console.error("Error:", error);
                    alert("Error");
                });
        }
    });
}
