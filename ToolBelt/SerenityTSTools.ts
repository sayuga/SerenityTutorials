module myTSTools {
    /**
     * Retuns Quickfilters object to fill search box quick filter
     * options while keepinglocalization.
     *
     * @param myRow any Row Type being referenced
     * @param fields string Array of field names. 
     */
    export function quickSearchList(myRow: any, fields: any[]): any[] {
        let txt = (s) => Q.text("Db." + myRow.localTextPrefix + "." + s).toLowerCase();
        var container = [];

        container.push({ name: "", title: "all" });
        
        if (fields.length>0){
            for (var x in fields) {
                container.push({ name: fields[x], title: txt(fields[x]) });
            }
        }

        return container;
    }        
    
      /**
         * categoryToggler will hide the targeted category group until the trigger field is given
         * a value. If the trigger field is empty, then the category will reset to being hidden.
         * @param triggerField form field used as trigger for behavior. Type LookupEditor
         * @param categoryName category group name to be targeted for behaviour. Type string
         */
      export function categoryToggler(triggerField: any, categoryName: string) {
            //  gets the parent element of the category field tha will exhibit behavior
            var ele = this.element.find(".category-title:contains('" + categoryName + "')").parent(); 

            //gets the value of the triggered field so that it can be checked
            var checkTrigger = triggerField.value;

            //  if the field value IS NOT initialy empty/null, it displays the field and adds the 
            //  behavior to the affected fields  
            if (Q.isEmptyOrNull(checkTrigger)) {
                ele.toggle(false);

                triggerField.changeSelect2(e => { //captures event change on trigger field
                    var myTrigger = triggerField.value;

                    //behavior when triggered. In my case, if the field is empty or null, it hides the category using toggle(true); 
                    if (Q.isEmptyOrNull(myTrigger) === true) {
                        ele.toggle(false); 
                    }

                    else {
                        ele.toggle(true);
                    }
                });
            }

            //  if the field value IS initially empty/null, it hides the field and adds the 
            //  behavior to the affected fields  
            else {
                triggerField.changeSelect2(e => {
                    var myTrigger = triggerField.value;

                    //behavior when triggered. In my case, if the field is empty or null, it hides the category using toggle(true); 
                    if (Q.isEmptyOrNull(myTrigger) === true) {
                        ele.toggle(false);
                    }

                    else {
                        ele.toggle(true);
                    }
                });
            }
        };
    
}

