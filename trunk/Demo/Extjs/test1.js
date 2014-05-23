Ext.onReady(function(){
	new Ext.Panel({
		renderTo: Ext.getBody(),
		title: "面板头部header",
		width: 400,
		height: 200,

		// 标题栏
		tools:[
			{id: "help", 
				handler: function(){Ext.Msg.alert('help','pleasehelpme!');}
			},
			{id: "save"},
			{id: "close"}
			],		

		// 面板内容区
		html: '<h1>面板主区域</h1>',

		// 顶部工具栏
		tbar:[
			{text:'顶部工具栏topToolbar'},
			{pressed: false, text: '刷新'},
			{xtype: 'tbfill'},
			{pressed: true, text:'添加'},
			{xtype: "tbseparator"},
			{pressed: true, text: '保存'}
			],

		// 底部工具栏
		bbar:[{
			text:'底部工具栏bottomToolbar'
			}],
		
		// 底部
		buttons:[{
			text:"按钮位于footer"
			}]
		});
	});