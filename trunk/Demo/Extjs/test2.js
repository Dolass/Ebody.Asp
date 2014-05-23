Ext.onReady(function(){
	new Ext.Viewport({
		enableTableScroll: true,
		layout: 'border',

		// 设定内部对像
		items: [
			{title: '面板',
			region: 'north',
			height: 50,
			html: '人事管理系统'},

			{title: '菜单',
			region: 'west',
			width: 200,
			collapsible: true,
			html: '菜单栏'},

			{xtype: 'tabpanel',
			region: 'center',
			items: [
				{title: '基本资料',
				html: '基本资料'},
				{title: '考勤记录',
				html: '考勤记录'}]
			}
		]		
	})
});
