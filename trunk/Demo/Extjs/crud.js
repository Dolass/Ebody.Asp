/*!
 * Ext JS Library 3.2.0
 * Copyright(c) 2006-2010 Ext JS, Inc.
 * licensing@extjs.com
 * http://www.extjs.com/license
 */
Ext.onReady(function(){
    var newFormWin;
    var userform;
    // create the Data Store  数据
    var store = new Ext.data.JsonStore({
        root: 'pageResult',
        totalProperty: 'totalRecords',
        idProperty: 'userId',
        remoteSort: true,
        fields: [
            'userEmail', 
            'userAge', 
            'userAddress', 
            'country',
            'userSex',
            'crtDate',
            'city',
            'province',
            'race',
            'userPhone',
            'userName',
            'userId'
        ],
        // load using script tags for cross domain, if the data in on the same domain as
        // this page, an HttpProxy would be better
        proxy: new Ext.data.HttpProxy({
            url: '请求数据路径'
        })
    });
    //分页显示用户数据
    var grid = new Ext.grid.GridPanel({
        autoFill : false,
        autoHeight : true,
        width:700,//宽度
        //height:300,//高度
        title:'用户信息表',
        store: store,//数据
        trackMouseOver:false,
        disableSelection:true,
        //加载的图标是否显示
        loadMask: true,
        //单行选中
        sm: new Ext.grid.CheckboxSelectionModel ({ singleSelect: false }),//复选框
        // grid columns列
        columns:[
        new Ext.grid.CheckboxSelectionModel({singleSelect : false}),//复选框
        //new Ext.grid.RowNumberer(),//序号
        {
            id: 'userName', // id assigned so we can apply custom css (e.g. .x-grid-col-topic b { color:#333 })
            header: "用户名",
            dataIndex: 'userName',
            //align: 'center',
            sortable: false
        },{
            header: "年龄",
            dataIndex: 'userAge',
            //align: 'center',
            sortable: false
        },{
            header: "email",
            dataIndex: 'userEmail',
            //align: 'center',
            sortable: false
        },{
            header: "电话",
            dataIndex: 'userPhone',
            //align: 'center',
            sortable: false
        }],
        // customize view config
        viewConfig: {
            forceFit:true,//True表示为自动展开/缩小列的宽度以适应grid的宽度，这样就不会出现水平的滚动条。
            enableRowBody:true,//True表示为在每一数据行的下方加入一个TR的元素
            showPreview:true,
            getRowClass : function(record, rowIndex, p, store){                 
                    return 'x-grid3-row-expanded';
                    // return 'x-grid3-row-collapsed';
            }
        },
        // 添加内陷的按钮
       tbar : [ {
            id : 'addUserForm',
            text : ' 新建  ',
            tooltip : '新建一个表单',
            iconCls : 'add',
            handler : function() {
                   add_btn();
            }
        }, '-', {
            id : 'editUserForm',
            text : '修改',
            tooltip : '修改',
            iconCls : 'edit',
            handler : function() {
                edit_btn();
            }
        }, '-', {
             text : '删除',
             tooltip : '删除被选择的内容',
             iconCls : 'remove',
             handler : function() {
                 handleDelete();
             }

        }],
        // paging bar on the bottom分页按钮
        bbar: new Ext.PagingToolbar({
            pageSize: 10,//每页条数
            store: store,//数据
            displayInfo: true,
            displayMsg: '从{0}条到{1}条  总共 {2}条',
            emptyMsg: "没有数据"
        })
    });
    // render it显示的层
    grid.render('user-grid');
    // trigger the data store load  加载用户
    store.load({params:{start:0, limit:10}});
    
    //添加用户按钮
     var add_btn = function() {
          addFormWin();
     };
     
      //添加和修改公用的用户的form=========开始========================================================================================================================== 
     var userForm = new Ext.FormPanel( {
            // collapsible : true,// 是否可以展开
            labelWidth : 75, // label settings here cascade unless overridden
            frame : true,
            bodyStyle : 'padding:5px 5px 0',
            waitMsgTarget : true,
            //reader : _jsonFormReader,
            defaults : {
                width : 230
            },
            defaultType : 'textfield',
            items : [{
                fieldLabel : 'id',
                name : 'userId',
                emptyText: 'id',
                hidden: true, 
                hideLabel:true,
                allowBlank : true
            }, {
                fieldLabel : '用户名',
                name : 'userName',
                emptyText: '用户名',
                allowBlank : false
            }, {
                fieldLabel : '年龄',
                name : 'userAge',
                emptyText: '年龄',
                  xtype : 'numberfield',
                allowBlank : false
            }, 
            new Ext.form.RadioGroup({
                fieldLabel : '性别',
                name:'userSex',
                items:[
                    {boxLabel: '男', name: 'userSex', inputValue: 1},
                        {boxLabel: '女', name: 'userSex', inputValue: 2}
                          ]  
            }), {
                fieldLabel : '种族',
                name : 'race',
                emptyText: '民族',
                allowBlank : false
            }, {
                fieldLabel : '电话',
                name : 'userPhone',
                emptyText: '联系电话',
                allowBlank : false
            }, {
                fieldLabel : 'Email',
                name : 'userEmail',
                vtype:'email',
                vtypeText:"不是有效的邮箱地址",
                  allowBlank : false
            }, {
                fieldLabel : '国家',
                name : 'country',
                emptyText: '国家',
                allowBlank : false
            },{
                fieldLabel : '省市',
                name : 'province',
                emptyText: '省市',
                allowBlank : false
            }, {
                fieldLabel : '城市',
                name : 'city',
                emptyText: '城市',
                allowBlank : false
            }, {
                fieldLabel : '地址',
                name : 'userAddress',
                emptyText: '地址',
                allowBlank : false
            }]          
        });
     //添加和修改公用的用户的form=========结束==========================================================================================================================   
        
     
        
    //添加操作开始========================================================================================================================== 
      // form_win定义一个Window对象，用于新建和修改时的弹出窗口。
     //添加用户的window
        var addFormWin = function() {
            // create the window on the first click and reuse on subsequent
            // clicks 判断此窗口是否已经打开了，防止重复打开
            if (!newFormWin) {
                newFormWin = new Ext.Window( {
                    el : 'topic-win',
                    layout : 'fit',
                    width : 400,
                    height : 400,
                    closeAction : 'hide',
                    plain : true,
                    title : '添加用户',
                    items : userForm,
                    buttons : [ {
                        text : '保存',
                        disabled : false,
                        handler :
                            addBtnsHandler
                        }, {
                        text : '取消',
                        handler : function() {
                            userForm.form.reset();//清空表单
                            newFormWin.hide();
                        }
                    }]
                });
            }
            newFormWin.show('addUserForm');//显示此窗口
        }
        //添加操作按钮
        function addBtnsHandler() {
            if (userForm.form.isValid()) {
                  userForm.form.submit( {
                      url : '请求数据路径', 
                      waitMsg : '正在保存数据，稍后...',
                      success : function(form, action) {
                                Ext.Msg.alert('保存成功', '添加用户信息成功！');
                                userForm.form.reset();//清空表单
                                grid.getStore().reload();
                                newFormWin.hide();
                      },
                      failure : function(form, action) {
                                  Ext.Msg.alert('保存失败', '添加人员信息失败！');
                      }
                  });
            }
            else {
                 Ext.Msg.alert('信息', '请填写完成再提交!');
            }                
        }
     //添加操作结束========================================================================================================================== 
        
     //修改操作开始==========================================================================================================================    
        //点击修改按钮加载数据   
        function edit_btn(){　　 
            var selectedKeys = grid.selModel.selections.keys;//returns array of selected rows ids only　　　　　　
            //判断是否选中一行数据 没有选中提示没有选中，选中加载信息
            if(selectedKeys.length != 1){　　　　　　　　
                Ext.MessageBox.alert('提示','请选择一条记录！');　　　　　　
                }　//加载数据　
                else{
                    var EditUserWin = new Ext.Window({　　　　　　　　
                    title: '修改员工资料',　//题头　　　　　　　
                    layout:'fit',//布局方式　　　　　　　　
                    width:400,//宽度　　　　　
                       height:400,//高度　　　　　　　　
                    plain: true,//渲染　　　　　　　　
                    items:userForm,　　　　　　　
                    //按钮
                    buttons: [{　　　　　　　　　　
                        text:'保存',　　　　　　　　　
                        handler:function(){
                            updateHandler(EditUserWin); 　
                        }　　　　　　
                    },{　　　　　　　　　　
                        text: '取消',　　　　　　　　　　
                        handler: function(){　　　　　　　　　　　　
                            EditUserWin.hide();　　　　　　　　　　
                        }　
                    }]　　 
                });
                EditUserWin.show("editUserForm");
                    loadUser();
                }
        }     
           //加载数据
           function loadUser(){
               var selectedKeys = grid.selModel.selections.keys;//returns array of selected rows ids only
               userForm.form.load({                    
                    waitMsg : '正在加载数据请稍后',//提示信息                
                    waitTitle : '提示',//标题                
                    url : '请求数据路径',            
                    params:{USER_ID:selectedKeys},                
                    method:'POST',//请求方式                            
                    failure:function(form,action){//加载失败的处理函数                    
                        Ext.Msg.alert('提示','数据加载失败');                
                    }            
               });        
           }　
         //修改按钮操作
           function updateHandler(w){
            if (userForm.form.isValid()) {
                userForm.form.submit({                    
                    clientValidation:true,//进行客户端验证                
                    waitMsg : '正在提交数据请稍后...',//提示信息                    
                    waitTitle : '提示',//标题                
                    url : 'http://localhost:8080/mypo/users/UserManagerAction/updateUser.json',//请求的url地址                    
                    method:'POST',//请求方式                    
                    success:function(form,action){//加载成功的处理函数    
                         w.hide();
                         userForm.form.reset();//清空表单
                         grid.getStore().reload();                    
                         Ext.Msg.alert('提示','修改信息成功');                    
                    },
                    failure:function(form,action){//加载失败的处理函数                        
                         Ext.Msg.alert('提示','ID不能修改');
                         Ext.Msg.alert('提示','修改信息失败');                    
                    }                
                });    
            }else {
                Ext.Msg.alert('信息', '请填写完成再提交!');  
            }
           }　　　
     //修改操作结束==========================================================================================================================  
           
           
    //删除操作开始==========================================================================================================================
      function handleDelete(){　　　
              var selectedKeys = grid.selModel.selections.keys; //returns array of selected rows ids only　　　　　　
               if(selectedKeys.length > 0)　　　　　　{　　　　　　　　
                Ext.MessageBox.confirm('提示','您确实要删除选定的记录吗？', deleteRecord);　　　　　　
            }else{　　　　　　　　
                Ext.MessageBox.alert('提示','请至少选择一条记录！');　　　　　　
               }//end　　
      }　
      //删除记录　　　
      function deleteRecord(btn){　　　　 
            if(btn=='yes'){　　　　　　　
                //var selectedRows = grid.selModel.selections.items;//returns record objects for selected rows (all info for row)　获得整行数据　　　　　　　
                    var selectedKeys = grid.selModel.selections.keys;//选中的行的值id
                Ext.MessageBox.show({　　　　　　　　　　　 
                    msg: '正在请求数据, 请稍侯',　　　　　　　　　　　 
                    progressText: '正在请求数据',　　　　　　　　　　　 
                    width:300,　　　　　　　　　　　 
                    wait:true,　　　　　　　　　　　 
                    waitConfig: {interval:200}　　　　　　　　 
                });　　　　　　　　
                Ext.Ajax.request({　　　　　　　　　　　　
                   url: '请求数据路径', //url to server side script　　　　　　　　　　　　
                   method: 'POST',　　　　　　　　　　　　
                   params:{USER_ID:selectedKeys},//the unique id(s)　　　　　　　　　　　　　　　　　　　　　　　
                   failure:function(){　　　　　　　　　　　　　　
                        Ext.MessageBox.hide();　　　　　　　　　　　　　　
                        Ext.MessageBox.alert("警告","出现异常错误！请联系管理员！");　　　　　　　　　　　　
                   },　　　　　　　　　　　　　
                   success:function(){　　　　　　　　　　　　　　
                        Ext.MessageBox.hide();
                        Ext.MessageBox.alert("成功","删除成功！");　　　　　　　　　　　　　　
                        store.reload();　　　　　　　　　　　　
                   }　　　　　　　　　　　　　　　　　　　　　　　　　
                })// end Ajax request
            }
        }
      //删除操作结束==========================================================================================================================
});