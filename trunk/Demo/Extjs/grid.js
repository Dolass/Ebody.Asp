Ext.require([
    'Ext.grid.*',
    'Ext.data.*',
    'Ext.selection.CheckboxModel',
	'Ext.selection.CellModel',
]);

Ext.onReady(function(){
	// 建立数据模板
    Ext.define('MyData', {
        extend: 'Ext.data.Model',
        fields: [
			{name: 'user_id', type: 'Auto', hidden: true, sortable: false},
            {name: 'user_name', type: 'Auto'},
            {name: 'password', type: 'Auto'},
            {name: 'desc', type: 'Auto'},
			{name: 'sex', type: 'Auto'},
			{name: 'born_date', type: 'date', dateFormat: 'Y/m/d'},
            {name: 'valid', type: 'Auto'}
         ]
    });

	// 显示提示
    Ext.QuickTips.init();

	// 创建数据源
    var store = Ext.create('Ext.data.Store', {
        pageSize: 5,	// 页面记录数
        model: 'MyData',
        remoteSort: false, // 当为false时，在本地进行当页数据排序，为true时，是后台进行全局排序
        proxy: {
            // load using script tags for cross domain, if the data in on the same domain as
            // this page, an HttpProxy would be better
            type: 'ajax',
            url: 'getdata.asp', // 请求数据页面
			method : 'post',
            reader: {
				type: 'json',
                root: 'items',
				totalProperty: 'total'
            },
            // sends single sort as multi parameter
            simpleSortMode: true
        },
        sorters: [{
            property: 'user_name',
            direction: 'DESC'
        }]
    });

	// 创建复选框对像
	var sm = Ext.create('Ext.selection.CheckboxModel');

	// 创建编辑对像
	var cellEditing = Ext.create('Ext.grid.plugin.CellEditing', {
        clicksToEdit: 2,
			clicksToMoveEditor: 1,  
                autoCancel: false
    });

    ////////////////////////////////////////////////////////////////////////////////////////
    // Grid 1
    ////////////////////////////////////////////////////////////////////////////////////////
    var grid1 = Ext.create('Ext.grid.Panel', {
        store: store,
        selModel: sm,
        columns: [
			// 自动行号
			Ext.create('Ext.grid.RowNumberer', {header: '行号', minWidth: 50}),

			// col1
            {text: "用户名", dataIndex: 'user_name', minWidth: 80,
				// 设有以下属性，则对像可编辑
				editor: {
					allowBlank: false	// 是否允许为空值
					}
				},

			// col2
            {text: "密码", dataIndex: 'password', width: 80},

			// col3
			// flex属性，用于栏位总是自动填充界面，这时width属性失效
            {text: "描述", dataIndex: 'desc', flex: 1, width: 200,
				editor: {
					allowBlank: false	// 是否允许为空值
					}	
			},

			// col4
            {text: "性别", dataIndex: 'sex', maxWidth: 50},

			// col5			
            {text: "出生年月", dataIndex: 'born_date', width: 120, renderer: Ext.util.Format.dateRenderer('m/d/Y'),				
				// 可编辑
				editor: {
					xtype: 'datefield',
					format: 'Y/m/d',
					minValue: '1900/01/01',
					disabledDays: [0, 6],
					disabledDaysText: 'Plants are not available on the weekends'
					}
			},

			// col6
            {text: "是否有效", dataIndex: 'valid',
				renderer: function(value){
				if(value==true){return '可用';}else{return '不可用';}
				}
			},
			
			// 删除控件
            {xtype: 'actioncolumn',
			text: '操作',
            width: 50,
            sortable: false,
            items: [{
                icon: '../../../../Resource/icons/fam/delete.gif',
                tooltip: 'Delete Row',
                handler: function(grid, rowIndex, colIndex) {
                    store.removeAt(rowIndex);
					}
				}]
			}

			],
		
        columnLines: true,
        width: 800,
        height: 300,
        frame: true,
		collapsible: true,
        animCollapse: true,
		resizable : true,
		floating: true,
		draggable: true,
        title: '用户信息',
        iconCls: 'icon-grid',

		viewConfig: {
			id: 'demo',
			loadMask:false,
			trackOver: true,
            stripeRows: true
        },

		// 头部控制条
		tbar: [{
            text: '新增数据',
            handler : function(){
                // Create a model instance
                var r = Ext.create('MyData', {
                    user_name: 'New user',
                    password: 0,
                    desc: '新用户'
                });
                store.insert(0, r);
                cellEditing.startEditByPosition({row: 0, column: 0});
            }
        }],
        plugins: [cellEditing],
		
		bbar: Ext.create('Ext.PagingToolbar', {
            store: store,
            displayInfo: true,
            displayMsg: '显示 {0} - {1} 条，共计 {2} 条',
            emptyMsg: '没有数据'
        }),
        renderTo: Ext.getBody()
    });


	// trigger the data store load
    store.loadPage(1);
});
