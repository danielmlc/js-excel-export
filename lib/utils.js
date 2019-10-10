/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 21:42:48
 * @LastEditTime: 2019-10-10 18:25:55
 */
function convertStr(value){
    return JSON.parse(value || '{}', function (k, v){
        if(v.indexOf && v.indexOf('function')>-1){
            return new Function( 'return ' + v )()
        }
        return v;
    })
  }

  module.exports = {
    convertStr
  }