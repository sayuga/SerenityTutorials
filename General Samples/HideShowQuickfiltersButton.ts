/// <reference path="../../../Northwind/Order/OrderGrid.ts" />

namespace SereneSandbox.BasicSamples {

    @Serenity.Decorators.registerClass()
    export class HideShowQuickFilters extends Northwind.OrderGrid {

        constructor(container: JQuery) {
            super(container);
        }

        getButtons() {
            var buttons = super.getButtons();


            buttons.push({
                separator: true,
                title: 'Quick Filters',
                cssClass: 'show-quick-filters',
                icon: 'fa-search',
                onClick: () => {

                    this.element.find('.quick-filters-bar').toggle(this.element.show)

                },
                hint: 'Click here to display Quick Filters'
            });

            return buttons;
        }


    }
}
