Ext.application({
	name:"HelloExt", 
	launch:function () {
		Ext.create('Ext.panel.Panel', { 
	        width: 1024, 
	        height: 720, 
	        layout: 'border', 
	        items: [{ 
	            region: 'south', 
	            xtype: 'panel', 
	            height: 20, 
	            split: false,
	            html: '欢迎登录!', 
	            margins: '0 5 5 5' 
	        },{ 
	            title: 'West Region is collapsible', 
	            region:'west', 
	            xtype: 'panel', 
	            margins: '5 0 0 5', 
	            width: 200, 
	            collapsible: true, 
	            id: 'west-region-container', 
	            layout: 'fit' 
	        },{ 
	            title: 'Center Region', 
	            region: 'center', 
	            xtype: 'panel', 
	            layout: 'fit', 
	            margins: '5 5 0 0', 
	            html:'在Extjs4中，center区域必须指定，否则会报错。' 
	        }], 
	        renderTo: Ext.getBody() 
    	});
	}
});
