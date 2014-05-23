Ext.onReady(function(){

	new Ext.Panel({
		renderTo: Ext.getBody(),
		width: 400,
		height: 200,
		layout: 'column',
		items: [
			{columnWidth: 0.5,
			title: '主管'},
			{columnWidth: 0.5,
			title: '员工',
				items:[{xtype: 'button', text: 'OK',handler: function(){Ext.Msg.alert('help','pleasehelpme!');}}]}
			]
		
	})
});
