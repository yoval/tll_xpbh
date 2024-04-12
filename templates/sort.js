function sortFiles() {
    const fileList = document.getElementById('file-list');
    const links = Array.from(fileList.getElementsByTagName('a'));

    links.sort((a, b) => {
        const aTime = a.textContent.split('_').pop();
        const bTime = b.textContent.split('_').pop();
        return bTime.localeCompare(aTime);
    });

    fileList.innerHTML = '';
    links.forEach(link => fileList.appendChild(link));
}

document.addEventListener('DOMContentLoaded', sortFiles);