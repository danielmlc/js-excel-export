/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 21:42:48
 * @LastEditTime: 2019-10-11 13:04:57
 */


function convertStr(value){
    return JSON.parse(value || '{}', function (k, v){
        if(v.indexOf && v.indexOf('function')>-1){
            return new Function( 'return ' + v )()
        }
        return v;
    })
  }
  
/**
 * @param target
 * @param exec 取值属性
 * @returns {*}
 */
function getter(target, exec = '_') {
    return new Proxy({}, {
      get: (o, n) => {
        return n === exec ?
          target :
          getter(typeof target === 'undefined' ? target : target[n], exec)
      }
    });
  }
  
  module.exports = {
    convertStr,
    getter
  }