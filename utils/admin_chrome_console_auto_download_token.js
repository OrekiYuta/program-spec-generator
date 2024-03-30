const accessToken = localStorage.getItem("Token");
const refreshToken = localStorage.getItem("refreshToken");

const jsonData = {
    "access_token": accessToken,
    "expires_in": 300,
    "refresh_expires_in": 1800,
    "refresh_token": refreshToken,
    "token_type": "Bearer",
    "id_token": "",
    "not-before-policy": 0,
    "session_state": "3ef78e15-ce3d-4ab9-b6b8-e19252a14c44",
    "scope": ""
};

const jsonString = JSON.stringify(jsonData);
const blob = new Blob([jsonString], {type: 'application/json'});
const a = document.createElement('a');

a.href = window.URL.createObjectURL(blob);
a.download = 'portal_user_token.json';
document.body.appendChild(a);
a.click();
document.body.removeChild(a);
