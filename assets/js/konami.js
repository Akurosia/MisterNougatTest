// a key map of allowed keys
var allowedKeys = {
  37: 'left',
  38: 'up',
  39: 'right',
  40: 'down',
  65: 'a',
  66: 'b'
};

// the 'official' Konami Code sequence
var konamiCode = ['up', 'up', 'down', 'down', 'left', 'right', 'left', 'right', 'b', 'a'];

// a variable to remember the 'position' the user has reached so far.
var konamiCodePosition = 0;

// add keydown event listener
document.addEventListener('keydown', function(e) {
  // get the value of the key code from the key map
  var key = allowedKeys[e.keyCode];
  // get the value of the required key from the konami code
  var requiredKey = konamiCode[konamiCodePosition];

  // compare the key with the required key
  if (key == requiredKey) {

    // move to the next key in the konami code sequence
    konamiCodePosition++;

    // if the last key is reached, activate cheats
    if (konamiCodePosition == konamiCode.length) {
      toogleKonami();
      konamiCodePosition = 0;
    }
  } else {
    konamiCodePosition = 0;
  }
});

function toogleKonami() {
  if (localStorage.getItem('konami_cookie') === "true") {
    localStorage.setItem('konami_cookie', false)
    document.getElementById("commandlist").style.display = "none";
  } else {
    //document.body.style.backgroundImage = "url('images/cheatBackground.png')";
    document.getElementById("commandlist").style.display = "";
    localStorage.setItem('konami_cookie', true)
  }
}

if (localStorage.getItem('konami_cookie') === "true") {
  document.getElementById("commandlist").style.display = "";
}
