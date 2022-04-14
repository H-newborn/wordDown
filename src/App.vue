
<template>
  <div id="app">
    <img alt="Vue logo" src="./assets/logo.png">
    <button @click="renderDoc">
      Render docx template
    </button>
  </div>
</template>

<script>
/* eslint-disable */
// 基本模块
import Docxtemplater from 'docxtemplater'
import PizZip from 'pizzip'
import PizZipUtils from 'pizzip/utils/index.js'
import { saveAs } from 'file-saver'

// 图片模块
import ImageModule from 'docxtemplater-image-module-free'


// 解析语法模块
import expressions from 'angular-expressions'
import assign from 'lodash/assign'

expressions.filters.lower = function (input) {
  // This condition should be used to make sure that if your input is
  // undefined, your output will be undefined as well and will not
  // throw an error
  if (!input) return input;
  return input.toLowerCase();
}

function angularParser(tag) {
  tag = tag
      .replace(/^\.$/, "this")
      .replace(/('|')/g, "'")
      .replace(/("|")/g, '"');
  const expr = expressions.compile(tag);
  return {
    get: function (scope, context) {
      let obj = {};
      const scopeList = context.scopeList;
      const num = context.num;
      for (let i = 0, len = num + 1; i < len; i++) {
          obj = assign(obj, scopeList[i]);
      }
      return expr(scope, obj);
    },
  };
}
// 解析语法模块

// 加载文件
function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback)
}

// 配置空值替换函数 作为配置参数可配置在setOptions中
function nullGetter(part, scopeManager) {
  if (!part.module) {
    return "-null-";
  }
  if (part.module === "rawxml") {
    return "";
  }
  return "--";
}

export default {
  name: 'App',
  components: {
  },
  methods: {
    // 基础文字变量替换
    renderDoc() {
      // 本地word.docx文件需要放在public目录下
      loadFile('/word.docx',
        function(error, content) {
          if (error) {
            throw error
          }
          const zip = new PizZip(content)
          // 没有配置解析语法，深层次对象语法（obj.xx.xx）不可识别
          const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
          })

          doc.render({
            subject: '示例',
            title: '简介',
            hasKitty: true,
            kitty: "Minie",
            hasDog: false,
            dog: "wangwang",
            users: [1,2,3,4],
            user: {
              name: '张三',
              age: 15,
              num: {
                say: 'hello'
              }
            },
            loop:[
              { name: "Windows", price: 100 },
              { name: "Mac OSX", price: 200 },
              { name: "Ubuntu", price: 0 }
            ],
            userGreeting: (scope) => {
              return "The product is" + scope.name + ", price：" + scope.price;
            },
          })
          
          const out = doc.getZip().generate({
            type: 'blob',
            mimeType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          })
          // Output the document using Data-URI
          saveAs(out, 'outputDoc.docx')
        }
      )
    },
    // 文字变量+图片
    renderImg() {
      // 本地word.docx文件需要放在public目录下
      loadFile('/word.docx',
      function(error, content) {
        if (error) {
          throw error
        }

        // 图片配置
        const imageOpts = {
          getImage: function(tagValue, tagName) {
            return new Promise(function (resolve, reject) {
              PizZipUtils.getBinaryContent(tagValue, function (error, content) {
                if (error) {
                  return reject(error);
                }
                return resolve(content);
              });
            });
          },
          getSize : function (img, tagValue, tagName) {
            // FOR FIXED SIZE IMAGE :
            return [150, 150];
          }
        }

        var imageModule = new ImageModule(imageOpts);

        const zip = new PizZip(content)

        // 实例化有两种方式 这里是链式
        const doc = new Docxtemplater().loadZip(zip).setOptions({
          // delimiters: { start: "[[", end: "]]" },
          paragraphLoop: true,
          linebreaks: true,
          nullGetter: nullGetter,
          parser: angularParser
        }).attachModule(imageModule).compile()

        doc.renderAsync({
          // 图片路径
          image: '/logo.png',
          subject: '示例',
          title: '简介',
          hasKitty: true,
          kitty: "Minie",
          hasDog: false,
          dog: "wangwang",
          users: [1,2,3,4],
          user: {
            name: '张三',
            age: 15,
            num: {
              say: 'hello'
            }
          },
          loop:[
            { name: "Windows", price: 100 },
            { name: "Mac OSX", price: 200 },
            { name: "Ubuntu", price: 0 }
          ],
          userGreeting: (scope) => {
              return "The product is" + scope.name + ", price：" + scope.price;
          },
        }).then(function () {
          const out = doc.getZip().generate({
              type: "blob",
              mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          });
          saveAs(out, "outputImg.docx");
        });
      })
    }
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
