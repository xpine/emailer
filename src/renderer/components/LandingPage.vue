<template>
  <div id="wrapper">
    <el-form ref="form" :model="form" :rules="rules" style="width:500px" label-width="100px">
      <el-form-item label="姓名：" prop="username">
        <el-input type="email" v-model="form.username" placeholder="请输入发送者姓名"></el-input>
      </el-form-item>
      <el-form-item label="邮箱账号：" prop="email">
        <el-input type="email" v-model="form.email" placeholder="请输入发送者邮箱账号"></el-input>
      </el-form-item>
      <el-form-item label="邮箱密码：" prop="password">
        <el-input type="password" v-model="form.password" placeholder="请输入密码"></el-input>
      </el-form-item>
      <el-form-item prop="test">
        <el-button @click="test">发送测试邮件</el-button>
      </el-form-item>
    </el-form>
    <el-form>
      <el-form-item label="选择excel文件：">
        <input type="file"  accept=".csv, application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change="selectFile"/>
      </el-form-item>
    </el-form>
    <el-tabs v-model="active" >
      <el-tab-pane v-for="(t,i) in tables" :key="i" :label="t.name" :name="t.name">
        <el-table :data="t.data" @selection-change="handleSelectionChange">
          <el-table-column
            type="selection"
            :selectable="checkSelected"
            width="55">
          </el-table-column>
          <el-table-column v-for="c in t.th" :key="c" :label="c" :prop="c"></el-table-column>
          <el-table-column label="发送状态" prop="done"></el-table-column>
        </el-table>
        <el-button :loading="loading"  style="margin-top:20px" @click="post(i)"> 发送邮件 </el-button>
      </el-tab-pane>
    </el-tabs>
  </div>  
</template>

<script>
  import {createTransporter} from './emailer.js';
  import {getCookie,setCookie} from './cookies.js';
  import fs from 'fs';
  import XLSX from 'xlsx';
  export default {
    name: 'landing-page',
    data(){
      return {
        loading:false,
        successEmail:[],
        errorEmail:[],
        active:null,
        selectRows :[],
        tables:[],
        rules:{
          username:[{
            required:true,
            message:'发送者姓名不能为空'
          }],
          email:[{
            required:true,
            message:'发送者邮箱不能为空'
          }],
          password:[{
            required:true,
            message:'发送者邮箱密码不能为空'
          }],
        },
        file:null,
        form:{
          username:getCookie('username'),
          email:getCookie('email'),
          password:getCookie('password'),
        },
      }
    },
    watch:{
      // 监听file改动，生成tables
      async file(file){
        var buf = await fs.readFileSync(file.path);
        var wb = XLSX.read(buf,{type:'buffer'});
        const {SheetNames,Sheets} = wb;
        const  tables = SheetNames.map(name=>{
          const data = XLSX.utils.sheet_to_json(wb.Sheets[name]);
          return {
            name:name,
            data:data,
            th:Object.keys(data[0]),
          }
        });
        this.tables = tables;
        this.active = tables[0] && tables[0].name || null;

      }
    },
    methods: {
      // 选择excel
      selectFile(e,file){
        this.file = e.target.files[0];
      },
      checkSelected(row,index){
        return !row.done
      },
      handleSelectionChange(val){
        this.selectRows = val;
        console.log(val,1)
      },
      // 开始发送
      post(index){
        const table = this.tables[index];
        const {
          name,
          data,
          th
        } = table;
        const {selectRows} = this;
        if(!selectRows.length){
          this.$message.error('请先选中');
          return 
        }
        this.loading = true;
        for(let i=0,max=selectRows.length;i<max;i++){
          // 生成mail options
          const item = selectRows[i];
          let mailOptions = {
            from: `${getCookie('username')} <${getCookie('email')}>`,
            to:item['邮箱'],
            subject: '工资',
            html:this.generateHtml(th,item)
          }
          // 发送邮件
          this.transporter.sendMail(mailOptions,(error,info)=>{
            if(error){
              this.errorEmail.push(error);
              return console.log(error);
            }
            this.successEmail.push(info);
            if((this.errorEmail.length + this.successEmail.length) == max ){
              this.$notify({
                title: '发送完毕',
                message: `共${max}条，成功${this.successEmail.length}条,失败${this.errorEmail.length}条`,
                type: 'success'
              });
              this.loading = false;
              this.successEmail = [];
              this.errorEmail = [];
            }
            const tables = JSON.parse(JSON.stringify(this.tables));
            // 如果selectrows 下存在__rowNum__ 属性使用它，如果不存在遍历比对email地址，确定发送成功的邮箱
            if(item.__rowNum__){
              tables[index].data[item.__rowNum__-1].done = '发送成功';
            }else{
              for(let j=0,max= data.length;j<max;j++){
                if(data[j]['邮箱'] == item['邮箱']){
                  tables[index].data[j].done = '发送成功';
                }
              }
            }
            this.tables = tables;
          })
        }
      },
      // 生成table html email content
      generateHtml(th,item){
        let html = ['<table>','</table>'];
        let row = ['<tr>','</tr>'];
        let header = row.join(th.reduce((sum,d)=>{
          return sum + '<td>' + d + '</td>';
        },''));
        let data = row.join(th.reduce((sum,d)=>{
          return sum + '<td>' + item[d] + '</td>';
        },''));
        html = html.join(header+data);
        return html;
      },
      // 测试完成之后将用户信息保存至cookie中方便下次使用
      test(){
        const {
          username,
          email,
          password
        } = this.form;
        this.$refs.form.validate((ok)=>{
          if(ok){
            this.transporter = createTransporter(email,password);
            const emailAccount = this.transporter.options.auth.user;
            let mailOptions = {
              from: `${username} <${emailAccount}>`,
              to: `${emailAccount}`, 
              subject: '测试邮件',
              html: '<b>测试邮件，请勿回复！</b>'
            };
            this.transporter.sendMail(mailOptions, (error, info) => {
              if (error) {
                this.$notify({
                  title: '发送失败',
                  message: error.message,
                  type: 'error'
                });
                return ;
              }
              this.$notify({
                title: '已发送',
                message: '请前往邮箱查看是否存在测试邮件',
                type: 'success'
              });
              setCookie('username',username);
              setCookie('email',email);
              setCookie('password',password);
            });
          }
        })
        
      },
    },
    mounted(){
      // 加载完成 判断是否已经测试成功过，如果是，创建transporter
      const {
        username,
        email,
        password
      } = this.form;
      if(username && email && password){
        this.transporter = createTransporter(email,password);
      }
    }
  }
</script>

<style>
  @import url('https://fonts.googleapis.com/css?family=Source+Sans+Pro');

  * {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
  }

  body { font-family: 'Source Sans Pro', sans-serif; }

  #wrapper {
    background:
      radial-gradient(
        ellipse at top left,
        rgba(255, 255, 255, 1) 40%,
        rgba(229, 229, 229, .9) 100%
      );
    height: 100vh;
    padding: 60px 80px;
    width: 100vw;
    overflow: auto;
  }

  #logo {
    height: auto;
    margin-bottom: 20px;
    width: 420px;
  }

  main {
    display: flex;
    justify-content: space-between;
  }

  main > div { flex-basis: 50%; }

  .left-side {
    display: flex;
    flex-direction: column;
  }

  .welcome {
    color: #555;
    font-size: 23px;
    margin-bottom: 10px;
  }

  .title {
    color: #2c3e50;
    font-size: 20px;
    font-weight: bold;
    margin-bottom: 6px;
  }

  .title.alt {
    font-size: 18px;
    margin-bottom: 10px;
  }

  .doc p {
    color: black;
    margin-bottom: 10px;
  }

  .doc button {
    font-size: .8em;
    cursor: pointer;
    outline: none;
    padding: 0.75em 2em;
    border-radius: 2em;
    display: inline-block;
    color: #fff;
    background-color: #4fc08d;
    transition: all 0.15s ease;
    box-sizing: border-box;
    border: 1px solid #4fc08d;
  }

  .doc button.alt {
    color: #42b983;
    background-color: transparent;
  }
</style>
