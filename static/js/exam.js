document.addEventListener('DOMContentLoaded', () => {
    const timerElement = document.getElementById('timer');
    const examForm = document.getElementById('examForm');

    if (!timerElement || !examForm) {
        return;
    }

    const endTimeText = timerElement.dataset.endTime;
    const endTime = new Date(endTimeText);

    const updateTimer = () => {
        const now = new Date();
        const remaining = Math.max(0, Math.floor((endTime.getTime() - now.getTime()) / 1000));
        const minutes = String(Math.floor(remaining / 60)).padStart(2, '0');
        const seconds = String(remaining % 60).padStart(2, '0');
        timerElement.textContent = `Time Left: ${minutes}:${seconds}`;

        if (remaining <= 0) {
            clearInterval(intervalId);
            examForm.submit();
        }
    };

    updateTimer();
    const intervalId = setInterval(updateTimer, 1000);
});
