/**
 * Autocompleter
 *
 * @license		MIT-style license
 * @author		Harald Kirschner <mail [at] digitarald.de>
 * @copyright	Author
 */
var Autocompleter = {};

Autocompleter.Base = new Class({

	options: {
		minLength: 1,
		useSelection: true,
		markQuery: true,
		inheritWidth: true,
		maxChoices: 10,
		injectChoice: null,
		onSelect: Class.empty,
		onShow: Class.empty,
		onHide: Class.empty,
		customTarget: null,
		className: 'autocompleter-choices',
		zIndex: 42,
		observerOptions: {},
		fxOptions: {},
		overflown: []
	},

	initialize: function(el, options) {
		this.setOptions(options);
		this.element = $(el);
		this.build();
		this.observer = new Observer(this.element, this.prefetch.bind(this), $merge({
			delay: 400
		}, this.options.observerOptions));
		this.value = this.observer.value;
		this.queryValue = null;
	},

	/**
	 * build - Initialize DOM
	 *
	 * Builds the html structure for choices and appends the events to the element.
	 * Override this function to modify the html generation.
	 */
	build: function() {
		if ($(this.options.customTarget)) this.choices = this.options.customTarget;
		else {
			this.choices = new Element('ul', {
				'class': this.options.className,
				styles: {zIndex: this.options.zIndex}
			}).injectInside(document.body);
			this.fix = new OverlayFix(this.choices);
		}
		this.fx = this.choices.effect('opacity', $merge({
			wait: false,
			duration: 200
		}, this.options.fxOptions))
			.addEvent('onStart', function() {
				if (this.fx.now) return;
				this.choices.setStyle('display', '');
				this.fix.show();
			}.bind(this))
			.addEvent('onComplete', function() {
				if (this.fx.now) return;
				this.choices.setStyle('display', 'none');
				this.fix.hide();
			}.bind(this)).set(0);
		this.element.setProperty('autocomplete', 'off')
			.addEvent(window.ie ? 'keydown' : 'keypress', this.onCommand.bindWithEvent(this))
			.addEvent('mousedown', this.onCommand.bindWithEvent(this, [true]))
			.addEvent('focus', this.toggleFocus.bind(this, [true]))
			.addEvent('blur', this.toggleFocus.bind(this, [false]))
			.addEvent('trash', this.destroy.bind(this));
	},

	destroy: function() {
		this.choices.remove();
	},

	toggleFocus: function(state) {
		this.focussed = state;
		if (!state) this.hideChoices();
	},

	onCommand: function(e, mouse) {
		if (mouse && this.focussed) this.prefetch();
		if (e.key && !e.shift) switch (e.key) {
			case 'enter':
				if (this.selected && this.visible) {
					this.choiceSelect(this.selected);
					e.stop();
				} return;
			case 'up': case 'down':
				if (this.queryValue == null) break;
				else if (!this.visible) this.showChoices();
				else {
					this.choiceOver((e.key == 'up')
						? this.selected.getPrevious() || this.choices.getLast()
						: this.selected.getNext() || this.choices.getFirst() );
					this.setSelection();
				}
				e.stop(); return;
			case 'esc': 
				e.stop(); this.hideChoices(); return;
		}
		this.value = false;
	},

	setSelection: function() {},

	hideChoices: function() {
		if (this.choices.style.left.toString().length < 1)
		{
			var pos = this.element.getCoordinates(this.options.overflown);
			this.choices.setStyles({
				left: pos.left,
				top: pos.bottom
			});
		}

		if (!this.visible) return;
		this.visible = this.value = false;
		this.observer.clear();
		this.fx.start(0);
		this.fireEvent('onHide', [this.element, this.choices]);
	},

	showChoices: function() {
		if (this.visible || !this.choices.getFirst()) return;
		if (this.element.value.length < 1) return;
		var echr = this.element.value.substr(this.element.value.length - 1);
		if (echr == "," || echr == ";") return;
		this.visible = true;
		var pos = this.element.getCoordinates(this.options.overflown);
		this.choices.setStyles({
			left: pos.left,
			top: pos.bottom
		});
		if (this.options.inheritWidth) this.choices.setStyle('width', pos.width);
		this.fx.start(1);
		this.choiceOver(this.choices.getFirst());
		this.fireEvent('onShow', [this.element, this.choices]);
	},

	prefetch: function() {
		var echr = this.element.value.substr(this.element.value.length - 1);
		if (echr == "," || echr == ";") {this.hideChoices(); return;}
		if (this.element.value.length < this.options.minLength) this.hideChoices();
		else if (this.element.value == this.queryValue) this.showChoices();
		else this.query();
	},

	updateChoices: function(choices) {
		this.choices.empty();
		this.selected = null;
		if (!choices || !choices.length) return;
		if (this.options.maxChoices < choices.length) choices.length = this.options.maxChoices;
		choices.each(this.options.injectChoice || function(choice, i){
			var el = new Element('li').setHTML(this.markQueryValue(choice));
			el.inputValue = choice;
			this.addChoiceEvents(el).injectInside(this.choices);
		}, this);
		this.showChoices();
	},

	choiceOver: function(el) {
		if (this.selected) this.selected.removeClass('autocompleter-selected');
		this.selected = el.addClass('autocompleter-selected');
	},

	choiceSelect: function(el) {
		var fsi = this.element.value.lastIndexOf(";");
		var dsi = this.element.value.lastIndexOf(",");
		var esi = Math.max(fsi, dsi);
		var inps = el.inputValue.substring(0, el.inputValue.length - 1);
		this.observer.value = this.element.value = this.element.value.substr(0, esi + 1) + inps + ",";
		this.hideChoices();
		this.fireEvent('onSelect', [this.element], 20);
	},

	/**
	 * markQueryValue
	 *
	 * Marks the queried word in the given string with <span class="autocompleter-queried">*</span>
	 * Call this i.e. from your custom parseChoices, same for addChoiceEvents
	 *
	 * @param		{String} Text
	 * @return		{String} Text
	 */
	markQueryValue: function(txt) {
		return (this.options.markQuery && this.queryValue) ? txt.replace(new RegExp('^(' + this.queryValue.escapeRegExp() + ')', 'i'), '<span class="autocompleter-queried">$1</span>') : txt;
	},

	/**
	 * addChoiceEvents
	 *
	 * Appends the needed event handlers for a choice-entry to the given element.
	 *
	 * @param		{Element} Choice entry
	 * @return		{Element} Choice entry
	 */
	addChoiceEvents: function(el) {
		return el.addEvents({
			mouseover: this.choiceOver.bind(this, el),
			mousedown: this.choiceSelect.bind(this, el)
		});
	}
});

Autocompleter.Base.implement(new Events);
Autocompleter.Base.implement(new Options);

Autocompleter.Local = Autocompleter.Base.extend({

	options: {
		minLength: 0,
		filterTokens : null
	},

	initialize: function(el, tokens, options) {
		this.parent(el, options);
		this.tokens = tokens;
		if (this.options.filterTokens) this.filterTokens = this.options.filterTokens.bind(this);
	},

	query: function() {
		this.hideChoices();
		var fsi = this.element.value.lastIndexOf(";");
		var dsi = this.element.value.lastIndexOf(",");
		var esi = Math.max(fsi, dsi);
		this.queryValue = this.element.value.substr(esi + 1);
		this.updateChoices(this.filterTokens());
	},

	filterTokens: function(token) {
		var regex = new RegExp('^' + this.queryValue.escapeRegExp(), 'i');
		return this.tokens.filter(function(token) {
			return regex.test(token);
		});
	}

});

var OverlayFix = new Class({

	initialize: function(el) {
		this.element = $(el);
		if (window.ie){
			this.element.addEvent('trash', this.destroy.bind(this));
			this.fix = new Element('iframe', {
				properties: {
					frameborder: '0',
					scrolling: 'no',
					src: 'javascript:false;'
				},
				styles: {
					position: 'absolute',
					border: 'none',
					display: 'none',
					filter: 'progid:DXImageTransform.Microsoft.Alpha(opacity=0)'
				}
			}).injectAfter(this.element);
		}
	},

	show: function() {
		if (this.fix) this.fix.setStyles($extend(
			this.element.getCoordinates(), {
				display: '',
				zIndex: (this.element.getStyle('zIndex') || 1) - 1
			}));
		return this;
	},

	hide: function() {
		if (this.fix) this.fix.setStyle('display', 'none');
		return this;
	},

	destroy: function() {
		this.fix.remove();
	}

});