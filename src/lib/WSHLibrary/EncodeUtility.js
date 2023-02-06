/*!
 * EncodeUtility.js v3
 *
 * Copyright (c) 2018 toshi (https://github.com/k08045kk)
 *
 * Released under the MIT license.
 * see https://opensource.org/licenses/MIT
 */

/**
 * WSH(JScript)用変換関連ライブラリ
 * @requires    module:ActiveXObject('MSXML.DOMDocument')
 * @requires    module:ActiveXObject('System.Text.UTF8Encoding')
 * @requires    module:ActiveXObject('System.Security.Cryptography.MD5CryptoServiceProvider')
 * @requires    module:ActiveXObject('System.Security.Cryptography.SHA1CryptoServiceProvider')
 * @requires    module:ActiveXObject('System.Security.Cryptography.SHA256Managed')
 * @requires    module:ActiveXObject('System.Security.Cryptography.RijndaelManaged')
 * @requires    FileUtility.js
 * @auther      toshi (https://github.com/k08045kk)
 * @version     3
 * @see         1.20180215 - add - 初版
 * @see         2.20180319 - update - PrivateUnderscore.jsを導入 - ソフトに共通処理化するため
 * @see         3.20180503 - update - リファクタリング
 */
(function(root, factory) {
    if (!root.EncodeUtility) {
      root.EncodeUtility = factory(root.FileUtility);
    }
  })(this, function(FileUtility) {
    "use strict";
    
    var _this   = void 0;
    var _doc    = _Process_createDOMDocument();
    var _encode = new ActiveXObject('System.Text.UTF8Encoding');
    // 補足:一般的にUTF8が標準のため、UTF-8とする。
    
    /**
     * PrivateUnderscore.js
     * @version   1
     */
    {
      function _isString(obj) {
        return Object.prototype.toString.call(obj) === '[object String]';
      };
      function _Process_createActiveXObjects(progIDs) {
        for (var i=0; i<progIDs.length; ++i) {
          try {
            return new ActiveXObject(progIDs[i]);
          } catch (e) {
            if (i == progIDs.length - 1) {  throw e;  }
          }
        }
        return null;
      };
      function _Process_createDOMDocument() {
        return _Process_createActiveXObjects([
          'MSXML2.DOMDocument.6.0',
          'MSXML2.DOMDocument.3.0',
          'Msxml2.DOMDocument',
          'Msxml.DOMDocument',
          'Microsoft.XMLDOM']);
      };
    }
    
    /**
     * コンストラクタ
     * @constructor
     */
    _this = function EncodeUtility_constrcutor() {};
    
    // Encodingを変更する
    // ASCIIEncoding, UTF7Encoding, UTF8Encoding, UnicodeEncoding, UTF32Encoding
    _this.setEncoding = function EncodeUtility_setEncoding(type) {
      var ret = false;
      try {
        var e = new ActiveXObject('System.Text.'+type);
        if (e != null) {
          _encode = null;
          _encode = e;
          ret = true;
        }
      } catch (e) {}
      return ret;
    };
    
    // 文字列→値
    function text2value(type, text) {
      var element = _doc.createElement('temp');
      element.dataType = type;            // 種別を設定
      element.text = text;                // 変換元を設定
      var ret = element.nodeTypedValue;   // 値を取り出す
      element = null;
      return ret;
    };
    
    // 値→文字列
    function value2text(type, value) {
      var element = _doc.createElement('temp');
      element.dataType = type;            // 種別を設定
      element.nodeTypedValue = value;     // 変換元を設定
      var ret = element.text;             // 値を取り出す
      element = null;
      return ret;
    };
    
    /**
     * 「文字列」から「byte配列」へ変換
     * @param {string} string - 元データ
     * @return {Byte[]} 変換データ
     */
    _this.str2bin = function EncodeUtility_string2bytes(string) {
      return _encode.GetBytes_4(string);
    };
    
    /**
     * 「byte配列」から「文字列」へ変換
     * @param {Byte[]} bytes - 元データ
     * @return {string} 変換データ
     */
    _this.bin2str = function EncodeUtility_bytes2string(bytes) {
      return _encode.GetString(bytes);
    };
    
    /**
     * 「byte配列」から「16進数文字列」
     * @param {Byte[]} bytes - 元データ
     * @return {string} 変換データ
     */
    _this.bin2hex = function EncodeUtility_bytes2hex(bytes) {
      return value2text('bin.hex', bytes);
    }
    
    /**
     * 「16進数文字列」から「byte配列」
     * @param {string} hex - 元データ
     * @return {Byte[]} 変換データ
     */
    _this.hex2bin = function EncodeUtility_hex2bytes(hex) {
      return text2value('bin.hex', hex);
    }
    
    /**
     * 「byte配列」から「Base64文字列」
     * @param {Byte[]} bytes - 元データ
     * @return {string} 変換データ
     */
    _this.bin2base64 = function EncodeUtility_bytes2base64(bytes) {
      return value2text('bin.base64', bytes);
    }
    
    /**
     * 「Base64文字列」から「byte配列」
     * @param {string} base64 - 元データ
     * @return {Byte[]} 変換データ
     */
    _this.base642bin = function EncodeUtility_base642bytes(base64) {
      return text2value('bin.base64', base64);
    }
    
    // ハッシュ化
    function hash(objname, bytes, opt_type) {
      var provider = new ActiveXObject(objname);
      
      if (_isString(bytes)) {
        // bytesが文字列(パス)指定の場合
        if (FileUtility && FileUtility.loadBinary) {
          bytes = FileUtility.loadBinary(bytes);
        } else {
          return null;
        }
      }
      if (bytes == null) {                // 値が未取得の時(ファイルなしor空ファイル)
        bytes = _this.str2bin('');        // 空のバイト配列を設定
      }
      provider.ComputeHash_2(bytes);
      var hashs = provider.Hash;
      provider.Clear();
      provider = null;                    // ActiveXObject開放
      bytes = null;                       // ファイルデータ開放
      
      // 結果の変換
      var ret  = hashs;
      var type = _isString(opt_type)? opt_type.toLowerCase(): 'hex';
      switch (type) {
      case 'str':
      case 'string':
        ret = _this.bin2str(hashs);
        break;
      case 'base64':
        ret = _this.bin2base64(hashs);
        break;
      case 'hex':
      default:
        ret = _this.bin2hex(hashs);
        break;
      }
      hashs = null;
      return ret;
    };
    
    /**
     * md5
     * @param {(Byte[]|string)} bytes - 元データ(string:ファイルパス)
     * @param {string} [opt_type='hex'] - 結果種別('bin'/'str'/'string'/'hex'/'base64')
     * @return {(string|Byte[]|unll)} 暗号データ
     */
    _this.md5 = function EncodeUtility_md5(bytes, opt_type) {
      return hash('System.Security.Cryptography.MD5CryptoServiceProvider', bytes, opt_type);
    };
    
    /**
     * sha1
     * @param {Byte[]|string} bytes - 元データ(string:ファイルパス)
     * @param {string} [opt_type='hex'] - 結果種別('bin'/'str'/'string'/'hex'/'base64')
     * @return {(string|Byte[]|unll)} 暗号データ
     */
    _this.sha1 = function EncodeUtility_sha1(bytes, opt_type) {
      return hash('System.Security.Cryptography.SHA1CryptoServiceProvider', bytes, opt_type);
    };
    
    // 暗号化/復号(Rijndael暗号)
    // http://qiita.com/tnakagawa/items/bba972e71cdc4f0f29f5
    function crypto(enc, key, src) {
      var crypt  = new ActiveXObject('System.Security.Cryptography.RijndaelManaged');
      
      // 共有キー設定
      var keyHash  = new ActiveXObject('System.Security.Cryptography.SHA256Managed');
      keyHash.ComputeHash_2(key);
      crypt.Key = keyHash.Hash;
      keyHash.Clear();
      keyHash = null;
      
      // 初期化ベクター設定
      var ivHash = new ActiveXObject('System.Security.Cryptography.MD5CryptoServiceProvider');
      ivHash.ComputeHash_2(key);
      crypt.IV = ivHash.Hash;
      ivHash.Clear();
      ivHash = null;
      
      // 暗号化/復号
      var dst = null;
      var len  = value2text('bin.hex', src).length / 2;
      var ts  = (enc)? crypt.CreateEncryptor(): crypt.CreateDecryptor();
      try {
        dst = ts.TransformFinalBlock(src, 0, len);
      } catch (e) {  // ブロック不正 or ...
      }
      ts = null;
      
      crypt.Clear();
      crypt = null;
      return dst;
    };
    
    /**
     * 暗号化(Rijndael暗号)
     * @param {Byte[]} key - 鍵
     * @param {Byte[]} src - 元データ
     * @return {Byte[]} 暗号データ
     */
    _this.encrypt = function EncodeUtility_encrypt(key, src) {
      return crypto(true, key, src);
    };
    
    /**
     * 復号(Rijndael暗号)
     * 復号データがnull以外でも、復号失敗の可能性がある。
     * @param {Byte[]} key - 鍵
     * @param {Byte[]} src - 暗号データ
     * @return {(Byte[]|null)} 復号データ(null:復号失敗)
     */
    _this.decrypt = function EncodeUtility_decrypt(key, src) {
      return crypto(false, key, src);
    };
    
    return _this;
  });
  