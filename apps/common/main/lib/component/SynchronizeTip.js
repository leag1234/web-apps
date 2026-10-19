/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
if (Common === undefined)
    var Common = {};

define([
    'common/main/lib/component/BaseView',
    'common/main/lib/core/NotificationCenter'
], function () {
    'use strict';

    Common.UI.SynchronizeTip = Common.UI.BaseView.extend(_.extend((function() {
        return {
            options : {
                target  : $(document.body),
                text    : '',
                placement: 'right-bottom',
                showLink: true,
                showButton: false,
                closable: true,
                textHeader: ''
            },

            template: _.template([
                '<div class="synch-tip-root <% if (!!scope.options.extCls) {print(scope.options.extCls + \" \");} %><%= scope.placement %>" style="<%= scope.style %>">',
                    '<div class="asc-synchronizetip">',
                        '<div class="tip-arrow <%= scope.placement %>"></div>',
                        '<div>',
                            '<div class="tip-text">',
                            '<% if ( scope.textHeader !== "" ) { %>',
                            '<label class="tip-header"><%= scope.textHeader %></label><br>',
                            '<% } %>',
                            '<%= scope.text %></div>',
                            '<% if ( scope.closable ) { %>',
                            '<div class="close"></div>',
                            '<% } %>',
                        '</div>',
                        '<% if ( scope.showLink ) { %>',
                        '<div class="show-link"><label><span><%= scope.textLink %></span></label></div>',
                        '<% } %>',
                        '<% if ( scope.showButton ) { %>',
                        '<div class="btn-div"><%= scope.textButton %></div>',
                        '<% } %>',
                    '</div>',
                '</div>'
            ].join('')),

            initialize : function(options) {
                const me = this;
                const app = (window.DE || window.PE || window.SSE || window.PDFE || window.VE);
                app.getController('Common.Controllers.Shortcuts').updateShortcutHints({
                    Save: {
                        label: '',
                        applyCallback: function(item, hintText) {
                            me.textSynchronize += hintText;
                        },
                        ignoreUpdates: true
                    },
                });
                
                Common.UI.BaseView.prototype.initialize.call(this, options);
                this.target = this.options.target;
                this.text = !_.isEmpty(this.options.text) ? this.options.text : this.textSynchronize;
                this.textLink = !_.isEmpty(this.options.textLink) ? this.options.textLink : this.textDontShow;
                this.placement = this.options.placement; // if placement='target' and position is undefined  show in top,left position of target, also use for arrow position
                this.showLink = this.options.showLink;
                this.showButton = this.options.showButton;
                this.closable = this.options.closable;
                this.textButton = this.options.textButton || this.textGotIt;
                this.textHeader = this.options.textHeader || '';
                this.position = this.options.position; // show in the position relative to target
                this.offset = this.options.offset; // shift from target
                this.style = this.options.style || '';
                this.automove = this.options.automove;
                this.binding = {};
            },

            render: function() {
                if (!this.cmpEl) {
                    this.cmpEl = $(this.template({ scope: this }));
                    $(document.body).append(this.cmpEl);
                    this.cmpEl.find('.close').on('click', _.bind(function() { this.trigger('closeclick');}, this));
                    this.cmpEl.find('.show-link label').on('click', _.bind(function() { this.trigger('dontshowclick');}, this));
                    this.cmpEl.find('.btn-div').on('click', _.bind(function() { this.trigger('buttonclick');}, this));

                    this.closable && this.cmpEl.addClass('closable');
                    this.binding.windowresize = _.bind(this.onWindowResize, this);
                }

                this.applyPlacement();

                return this;
            },

            show: function(){
                if (this.cmpEl) {
                    this.applyPlacement();
                    this.cmpEl.show()
                } else
                    this.render();
                this.automove && $(window).on('resize', this.binding.windowresize);
            },

            hide: function() {
                if (this.cmpEl) this.cmpEl.hide();
                this.trigger('hide');
                this.automove && $(window).off('resize', this.binding.windowresize);
            },

            close: function() {
                if (this.cmpEl) this.cmpEl.remove();
                this.trigger('close');
                this.automove && $(window).off('resize', this.binding.windowresize);
            },

            onWindowResize: function() {
                this.applyPlacement();
            },

            applyPlacement: function (repeatOnce) {
                var target = this.target && this.target.length>0 ? this.target : $(document.body);

                if (!target.is(':visible') && !repeatOnce) {
                    var me = this;
                    setTimeout(function(){ me.applyPlacement(true); }, 100);
                    return;
                }
                var showxy = Common.Utils.getOffset(target),
                    offset = this.offset || {x: 0, y: 0};
                if (this.placement=='target' && !this.position) {
                    this.cmpEl.css({top : showxy.top + 5  + offset.y + 'px', left: showxy.left + 5  + offset.x + 'px'});
                    return;
                }

                if (this.position && typeof this.position == 'object') {
                    var top = this.position.top, left = this.position.left, bottom = this.position.bottom, right = this.position.right;
                    if (bottom!==undefined || top!==undefined)
                        left = showxy.left + (target.width() - this.cmpEl.width())/2;
                    else
                        top = showxy.top + (target.height() - this.cmpEl.height())/2;
                    top = (top!==undefined) ? (top + 'px') : 'auto';
                    bottom = (bottom!==undefined) ? (bottom + 'px') : 'auto';
                    right = (right!==undefined) ? (right + 'px') : 'auto';
                    left = (left!==undefined) ? (left + 'px') : 'auto';

                    this.cmpEl.css({top : top, left: left, right: right, bottom: bottom});
                    return;
                }

                var placement = this.placement.split('-');
                if (placement.length>0) {
                    var top, left, bottom, right;
                    var pos = placement[0];
                    if (pos=='top') {
                        bottom = Common.Utils.innerHeight() - showxy.top + offset.y;
                    } else if (pos == 'bottom') {
                        top = showxy.top + target.height() + offset.y;
                    } else if (pos == 'left') {
                        right = Common.Utils.innerWidth() - showxy.left + offset.x;
                    } else if (pos == 'right') {
                        left = showxy.left + target.width() + offset.x;
                    }
                    pos = placement[1];
                    if (pos=='top') {
                        bottom = Common.Utils.innerHeight() - showxy.top - target.height()/2 + offset.y;
                    } else if (pos == 'bottom') {
                        top = showxy.top + target.height()/2 + offset.y;
                        var height = this.cmpEl.height();
                        if (top+height>Common.Utils.innerHeight())
                            top = Common.Utils.innerHeight() - height - 10;
                    } else if (pos == 'left') {
                        right = Common.Utils.innerWidth() - showxy.left - target.width()/2 + offset.x;
                    } else if (pos == 'right') {
                        left = showxy.left + target.width()/2 + offset.x;
                    } else {
                        if (bottom!==undefined || top!==undefined)
                            left = showxy.left + (target.width() - this.cmpEl.width())/2 + offset.x;
                        else
                            top = showxy.top + (target.height() - this.cmpEl.height())/2 + offset.y;
                    }
                    top = (top!==undefined) ? (top + 'px') : 'auto';
                    bottom = (bottom!==undefined) ? (bottom + 'px') : 'auto';
                    right = (right!==undefined) ? (right + 'px') : 'auto';
                    if (left!==undefined) {
                        var width = this.cmpEl.width();
                        if (left+width>Common.Utils.innerWidth())
                            left = Common.Utils.innerWidth() - width - 10;
                        if (left < 10)
                            left = 10;
  
                        left = (left + 'px');
                    } else
                        left = 'auto';
                    this.cmpEl.css({top : top, left: left, right: right, bottom: bottom});
                }
            },

            isVisible: function() {
                return this.cmpEl && this.cmpEl.is(':visible');
            },

            setText: function(text) {
                if (this.text !== text) {
                    this.text = text;
                    this.cmpEl.find('.tip-text').text(text);
                }
            },

            textDontShow        : 'Don\'t show this message again',
            textSynchronize     : 'The document has been changed by another user.<br>Please click to save your changes and reload the updates.',
            textGotIt: 'Got it'
        }
    })(), Common.UI.SynchronizeTip || {}));

    Common.UI.TooltipManager = new(function() {
        var _helpTips = {
            // 'step': {
            //     name: 'localstorage-id', // (or undefined when don't save option to localstorage) save 1 to localstorage to not show message again
            //     placement: 'bottom',
            //     position: '',
            //     offset: {x: 0, y: 0}
            //     text: '',
            //     header: '',
            //     target: '#id', // string or $el
            //     link: {text: 'link text', src: 'UsageInstructions\/....htm', url: 'www.example.com' }, // (or false) Open help page
            //     showButton: true, // true by default
            //     closable: true, // true by default
            //     callback: function() {} // call when close tip,
            //     next: '' // show next tooltip on close
            //     prev: '' // don't show tooltip if the prev was not shown
            //     automove: false // applyPlacement on window resize
            //     maxwidth: 250 // number or string '123px/none/...', 250 by default,
            //     extCls: '' //
            //     noHighlight: false // false by default,
            //     noArrow: false // false by default,
            //     multiple: false // false by default, show tip multiple times,
            //     isNewFeature: false // false by default, show "New" tip in the header
            // }
        },
        _targetStack = {
            // 'targetStr' : [targetEl1, targetEl2...]
        },
        _suspendedOwners = [],
        _pendingSteps = [],
        _resumeTimer = null;

        var _queueTip = function(step) {
            if (_pendingSteps.indexOf(step) < 0) _pendingSteps.push(step);
        };

        var _forgetPendingTip = function(step) {
            var index = _pendingSteps.indexOf(step);
            if (index >= 0) _pendingSteps.splice(index, 1);
        };

        var _resumePendingTips = function() {
            if (_suspendedOwners.length || Object.keys(_helpTips).some(function(step) {
                var props = _helpTips[step];
                return props && props.isNewFeature && props.tip && props.tip.isVisible();
            })) return;

            var pending = _pendingSteps.slice();
            for (var i = 0; i < pending.length; i++) {
                var step = pending[i];
                if (!_getNeedShow(step)) {
                    _forgetPendingTip(step);
                    continue;
                }
                // Resume one guide; its ordinary dismissal keeps any next-step chain first.
                _showTip(step);
                if (_helpTips[step] && _helpTips[step].tip && _helpTips[step].tip.isVisible()) return;
            }
        };

        var _suspend = function(owner) {
            if (_suspendedOwners.indexOf(owner) < 0) _suspendedOwners.push(owner);
            Object.keys(_helpTips).forEach(function(step) {
                var props = _helpTips[step];
                if (props && props.isNewFeature && props.tip && props.tip.isVisible()) {
                    _queueTip(step);
                    props.tip.cmpEl.addClass('elixir-guided-tip-suspended');
                    props.tip.target && props.tip.target.removeClass('highlight-tip');
                    props.tip.hide();
                }
            });
        };

        var _resume = function(owner, afterWindowEvent) {
            var index = _suspendedOwners.indexOf(owner);
            if (index < 0) return;
            _suspendedOwners.splice(index, 1);
            // Custom surfaces release ownership after removing their content.
            if (!afterWindowEvent) {
                _resumePendingTips();
                return;
            }
            // Window close/hide events precede removal of their target elements.
            if (!_suspendedOwners.length && _resumeTimer === null) {
                _resumeTimer = setTimeout(function() {
                    _resumeTimer = null;
                    _resumePendingTips();
                }, 0);
            }
        };

        var _addTips = function(arr) {
            for (var step in arr) {
                if (arr.hasOwnProperty(step) && !Common.localStorage.getItem(arr[step].name) && !(_helpTips[step] && _helpTips[step].tip && (_helpTips[step].tip.isVisible() || _helpTips[step].tip.cmpEl.hasClass('elixir-guided-tip-suspended')))) { // don't replace a visible or suspended tip
                    _helpTips[step] = arr[step];
                }
            }
        };

        var _removeTip = function(step) {
            _forgetPendingTip(step);
            var props = _helpTips[step];
            if (props && props.tip && props.tip.cmpEl.hasClass('elixir-guided-tip-suspended')) {
                props.tip.cmpEl.remove();
                if (typeof props.target === 'string' && props.stackIdx)
                    _targetStack[props.target][props.stackIdx-1] = undefined;
            }
            if (_helpTips[step]) {
                delete _helpTips[step];
                _helpTips[step] = undefined;
            }
        };

        var _getNeedShow = function(step) {
            return _helpTips[step] && !(_helpTips[step].name && Common.localStorage.getItem(_helpTips[step].name));
        };

        var _applyPlacement = function(step) {
            if (_helpTips[step] && _helpTips[step].tip && _helpTips[step].tip.isVisible())
                _helpTips[step].tip.applyPlacement();
        };

        var _closeTip = function(step, force, preventNext) {
            var steps = typeof step === 'string' ? [step] : step;
            steps && steps.forEach(function(step) {
                var props = _helpTips[step];
                if (props) {
                    preventNext && (props.next = undefined);
                    props.tip && props.tip.close();
                    props.tip = undefined;
                    force && props.name && Common.localStorage.setItem(props.name, 1);
                }
            });
        };

        var _findTarget = function(target) {
            if (typeof target === 'string') {
                if (!_targetStack[target])
                    _targetStack[target] = [];
                for (let i=_targetStack[target].length-1; i>=0; i--) {
                    if (_targetStack[target][i]) {
                        if (!_targetStack[target][i].hasClass('elixir-guided-tip-suspended'))
                            return _targetStack[target][i];
                    } else if (i === _targetStack[target].length - 1) {
                        _targetStack[target].pop();
                    }
                }
                return $(target);
            }
            return target;
        };

        var _showTip = function(step) {
            if (typeof step === 'object') { // init and show tip, object must have 'step' field
                if (step.step) {
                    if (!_helpTips[step.step])
                        _helpTips[step.step] = step;
                    else
                        _helpTips[step.step].text = step.text; // change text
                    step = step.step;
                }
            }
            if (!_helpTips[step]) return;
            if (_helpTips[step].isNewFeature && _suspendedOwners.length) {
                _queueTip(step);
                return false;
            }
            if (_getNeedShow(step) && !(_helpTips[step].prev && _getNeedShow(_helpTips[step].prev))) { // show current tip if previous tip has already been shown
                var props = _helpTips[step],
                    target = props.target,
                    targetEl = _findTarget(target);

                if (props.tip && props.tip.isVisible())
                    return true;

                if (!(targetEl && targetEl.length && targetEl.is(':visible')))
                    return false;

                if (props.tip && props.tip.cmpEl.hasClass('elixir-guided-tip-suspended')) {
                    if (!(props.tip.target && props.tip.target.is(':visible'))) return false;
                    props.tip.cmpEl.removeClass('elixir-guided-tip-suspended');
                    !props.noHighlight && props.tip.target.addClass('highlight-tip');
                    props.tip.show();
                    _forgetPendingTip(step);
                    return true;
                }

                var placement = props.placement || 'bottom';
                if (Common.UI.isRTL()) {
                    placement = placement.indexOf('right')>-1 ? placement.replace('right', 'left') : placement.replace('left', 'right');
                }
                !props.noHighlight && targetEl.addClass('highlight-tip');

                var header = props.header || '';
                if (props.isNewFeature && header) {
                    header = '<span>' + (Common.UI.SynchronizeTip.prototype.textNew || 'New') + '</span>' + header;
                }

                props.tip = new Common.UI.SynchronizeTip({
                    extCls: 'colored' + (props.extCls ? ' ' + props.extCls : '') + (props.noArrow ? ' no-arrow' : ''),
                    style: 'min-width:200px;max-width:' + (props.maxwidth ? props.maxwidth + (typeof props.maxwidth === 'number' ? 'px;' : ';') : '250px;'),
                    placement: placement,
                    position: props.position,
                    offset: props.offset,
                    target: targetEl,
                    text: props.text,
                    textHeader: header,
                    showLink: !!props.link,
                    textLink: props.link ? props.link.text : '',
                    closable: props.closable !== false, // true by default
                    showButton: props.showButton !== false, // true by default
                    textButton: props.textButton, // button text, Got it by default
                    automove: !!props.automove
                });
                props.tip.on({
                    'buttonclick': function() {
                        props.tip && props.tip.close();
                        props.tip = undefined;
                    },
                    'closeclick': function() {
                        props.tip && props.tip.close();
                        props.tip = undefined;
                    },
                    'dontshowclick': function() {
                        if (props.link.url)
                            window.open(props.link.url, '_blank');
                        else
                            Common.NotificationCenter.trigger('file:help', props.link.src);
                    },
                    'close': function() {
                        targetEl.removeClass('highlight-tip');
                        props.name && Common.localStorage.setItem(props.name, 1);
                        props.callback && props.callback();
                        props.next && _showTip(props.next);
                        !props.multiple && (delete _helpTips[step]);
                        if (typeof target === 'string' && props.stackIdx) {
                            _targetStack[target][props.stackIdx-1] = undefined;
                            props.stackIdx = undefined;
                        }
                        _forgetPendingTip(step);
                        _resumePendingTips();
                    }
                });
                props.tip.show();
                _forgetPendingTip(step);
                if (typeof target === 'string') {
                    _targetStack[target].push(props.tip.cmpEl);
                    props.stackIdx = _targetStack[target].length;
                }
            }
            return true;
        };

        // Only guided feature tips pause; SDK warnings retain their existing paths.
        Common.NotificationCenter.on({
            'modal:show': _suspend,
            'modal:close': function(owner) { _resume(owner, true); },
            'modal:hide': function(owner) { _resume(owner, true); }
        });

        return {
            suspend: _suspend,
            resume: _resume,
            showTip: _showTip,
            closeTip: _closeTip,
            removeTip: _removeTip,
            addTips: _addTips,
            getNeedShow: _getNeedShow,
            applyPlacement: _applyPlacement
        }
    })();
});

