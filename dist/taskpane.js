!function(){"use strict";var e={};function n(e,n,t,r,o,c,a){try{var u=e[c](a),i=u.value}catch(e){return void t(e)}u.done?n(i):Promise.resolve(i).then(r,o)}function t(e){return function(){var t=this,r=arguments;return new Promise((function(o,c){var a=e.apply(t,r);function u(e){n(a,o,c,u,i,"next",e)}function i(e){n(a,o,c,u,i,"throw",e)}u(void 0)}))}}function r(){return o.apply(this,arguments)}function o(){return o=t(regeneratorRuntime.mark((function e(){return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=t(regeneratorRuntime.mark((function e(n){var t;return regeneratorRuntime.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return(t=n.workbook.getSelectedRange()).load("address"),t.format.fill.color="yellow",e.next=5,n.sync();case 5:console.log("The range address was ".concat(t.address,"."));case 6:case"end":return e.stop()}}),e)})));return function(n){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),console.error(e.t0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),o.apply(this,arguments)}e.d=function(n,t){for(var r in t)e.o(t,r)&&!e.o(n,r)&&Object.defineProperty(n,r,{enumerable:!0,get:t[r]})},e.o=function(e,n){return Object.prototype.hasOwnProperty.call(e,n)},Office.onReady((function(){document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=r}))}();
//# sourceMappingURL=taskpane.js.map