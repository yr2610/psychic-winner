if (!Array.prototype.reduce) {
  Array.prototype.reduce = function reduce(accumulator){
    if (this===null || this===undefined) throw new TypeError("Object is null or undefined");

    var i = 0, l = this.length >> 0, curr;

    if(typeof accumulator !== "function") // ES5 : "If IsCallable(callbackfn) is false, throw a TypeError exception."
      throw new TypeError("First argument is not callable");

    if(arguments.length < 2) {
      if (l === 0) throw new TypeError("Array length is 0 and no second argument");
      curr = this[0];
      i = 1; // start accumulating at the second element
    }
    else
      curr = arguments[1];

    while (i < l) {
      if(i in this) curr = accumulator.call(undefined, curr, this[i], i, this);
      ++i;
    }

    return curr;
  };
}


// https://gist.github.com/jcxplorer/823878
function uuid() {
  var uuid = "", i, random;
  for (i = 0; i < 32; i++) {
    random = Math.random() * 16 | 0;

    if (i == 8 || i == 12 || i == 16 || i == 20) {
      uuid += "-"
    }
    uuid += (i == 12 ? 4 : (i == 16 ? (random & 3 | 8) : random)).toString(16);
  }
  return uuid;
}

//
// Stack (LIFO)
//
function Stack() { this.__a = new Array(); }
Stack.prototype.push = function(o) { this.__a.push(o); }
Stack.prototype.pop = function() { if( this.__a.length > 0 ) { return this.__a.pop(); } return null; }
Stack.prototype.peek = function() { if( this.__a.length > 0 ) { return this.__a[this.__a.length - 1]; } return null; }
Stack.prototype.size = function() { return this.__a.length; }
Stack.prototype.toString = function() { return '[' + this.__a.join(',') + ']'; }


// http://phiary.me/javascript-string-format/
// 文字列フォーマット(添字引数版)
// var str = "{0} : {1} + {2} = {3}".format("足し算", 8, 0.5, 8+0.5);
// 文字列フォーマット(オブジェクト版)
// str = "名前 : {name}, 年齢 : {age}".format( { "name":"山田", "age":128 } );
// 存在チェック
if (String.prototype.format == undefined) {  
  /**
   * フォーマット関数
   */
  String.prototype.format = function(arg)
  {
    // 置換ファンク
    var rep_fn = undefined;

    // オブジェクトの場合
    if (typeof arg == "object") {
      rep_fn = function(m, k) { return arg[k]; }
    }
    // 複数引数だった場合
    else {
      var args = arguments;
      rep_fn = function(m, k) { return args[ parseInt(k) ]; }
    }

    return this.replace( /\{(\w+)\}/g, rep_fn );
  }
}

// https://ja.wikipedia.org/wiki/Xorshift
var createXor128 = function(seed) {
  var x = 123456789;
  var y = 362436069;
  var z = 521288629;
  var w = (typeof seed === "undefined") ? 88675123 : seed;

  function random() {
    var t = x ^ (x << 11);
    x = y; y = z; z = w;
    w = (w ^ (w >>> 19)) ^ (t ^ (t >>> 8));
    return w / 0x100000000 + 0.5;
  }

  // 40回読み飛ばす
  // http://d.hatena.ne.jp/gintenlabo/20100928/1285702435 を鵜呑み
  for (var i = 0; i < 40; i++) {
    random();
  }

  return random;
}

/*
var seed = (new Date()).getTime();
//var seed = 201;
var random = createXor128(seed);
//var random = Math.random;
if (true) {
    var sRand = seed + ":\n";
    for (var i = 0; i < 10; i++) {
        //sRand += createRandomId(8, random) + "\n";
        sRand += random() + "\n";
    }
    WScript.Echo(sRand);
}

if (false) {
    var min = Math.pow(2, 53) - 1;
    var max = -(Math.pow(2, 53) - 1);
    for (var i = 0; i < 100000; i++) {
        var r = random();
        min = Math.min(r, min);
        max = Math.max(r, max);
    }
    WScript.Echo(min + "\n" + max);
}

*/
