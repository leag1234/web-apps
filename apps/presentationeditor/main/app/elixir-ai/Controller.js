/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define(['common/main/lib/elixir-ai/Controller','presentationeditor/main/app/elixir-ai/SlideTools','presentationeditor/main/app/elixir-ai/SlideStrings'], function (_controller,_tools,English) {
    'use strict';
    var AI=Common.ElixirAI,Base=Common.Controllers.ElixirAI;
    AI.SlideStrings=AI.SlideStrings||{};
    AI.slideText=function(key,values){var value=AI.SlideStrings[key]||English[key]||key;return value.replace(/\{(\w+)\}/g,function(m,k){return values&&values[k]!==undefined?String(values[k]):m;});};
    Common.Controllers.ElixirAI=Base.extend({
        toolsetId:'slide',
        initialize:function(){
            Base.prototype.initialize.apply(this,arguments);
            var me=this;
            this.addListeners({'DocumentHolder':{'createdelayedelements':function(view,type){if(type==='edit')me.installAskMenus(view);}}});
        },
        installAskMenus:function(view){
            if(!this.isAvailable())return;
            var me=this;view._elixirAIItems=view._elixirAIItems||{};
            ['slideMenu','textMenu','pictureMenu','tableMenu'].forEach(function(name){
                var menu=view[name];if(!menu||view._elixirAIItems[name])return;
                var item=new Common.UI.MenuItem({caption:AI.t('menuAsk'),iconCls:'elixir-ai-icon elixir-ai-icon-sparkle'});
                menu.addItem(item);item.on('click',function(){me.openBar();});view._elixirAIItems[name]=item;
            });
        }
    });
    return Common.Controllers.ElixirAI;
});
