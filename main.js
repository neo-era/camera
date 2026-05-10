// Quản lý danh sách camera treo trên trụ chiếu sáng
let cameras = [];

function renderCameras() {
    const list = document.getElementById('camera-list');
    list.innerHTML = cameras.map((cam, idx) => `
        <div>
            <b>Camera #${idx + 1}</b> - Vị trí: ${cam.location} - Trạng thái: ${cam.status}
        </div>
    `).join('');
}

function addCamera() {
    const location = prompt('Nhập vị trí trụ chiếu sáng:');
    if (location) {
        cameras.push({ location, status: 'Hoạt động' });
        renderCameras();
    }
}

window.onload = renderCameras;
