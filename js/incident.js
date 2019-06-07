/// <reference path="../../webresources/scripting/msxrmtools.xrm.page.js" />
/// <reference path="../../DEVELOP/webresources/ing_FetchSDK.js" />

if (typeof (Incdnt) == 'undefined')
{ Incdnt = {}; }

Incdnt = {
    ent_Subj: 'tisa_subject',
    atr_Prior: 'prioritycode',
    atr_Subj_Id: 'ing_subject_id',
    atr_Subsub_Id: 'ing_subsubj_id',
    atr_FirstResp: 'firstresponseslastatus',
    atr_SET_Sbjct: ['ing_casetypecode', 'ing_classifier_not_null', 'ing_classifier_visible', 'ing_office_not_null', 'ing_office_visible', 'ing_teamid', 'ing_time_count'],
    MAIN_defTeam: { id: '{50F9C1FD-DE2F-E911-810C-005056BA18B6}', name: 'Отдел по работе с обращениями' },
    MAIN_Config: null,
    svc_AttrValue: function (lName) {
        ///<summary>- FROM SLALOM Get the attribute value correctly or catch the null -</summary>
        ///<param name="lName" Type="String"> Name of the gettin attribute </param>
        var lAttrib = Xrm.Page.getAttribute(lName);

        return (typeof lAttrib == 'undefined' || lAttrib == null) ? null : lAttrib.getValue()
    },
    svc_AttribLookValue: function (lName) {
        ///<summary>- FROM SLALOM Get the value (lookup object) of lookup attribute correctly or catch the null -</summary>
        ///<param name="lName" Type="String"> Name of the gettin attribute </param>
        var lAttribValue = Incdnt.svc_AttrValue(lName);

        return (lAttribValue && 'length' in lAttribValue && lAttribValue.length > 0 && 'id' in lAttribValue[0]) ? lAttribValue[0] : null
    },
    svc_UpdateValue: function (lName, lValue, checkEmpty) {
        // Set the attribute value correctly
        var lAttrib = Xrm.Page.getAttribute(lName);

        if (lAttrib) {
            var atrValue = lAttrib.getValue();

            if (checkEmpty && atrValue != null) return;

            if (atrValue != lValue) lAttrib.setValue(lValue);
        }
    },
    svc_CntrlVisible: function (lName, lVisible) {
        // Set the control visibility
        var lCntrl = Xrm.Page.getControl(lName);

        if (lCntrl) lCntrl.setVisible(lVisible);
    },
    svc_RemOptions: function (lName, remIndxs) {
        // Remove options by ids from 'remIndxs' array
        var lCntrl = Xrm.Page.getControl(lName);

        if (lCntrl)
            for (var stp = 0, sLen = remIndxs.length; stp < sLen; stp++)
                lCntrl.removeOption(remIndxs[stp]); // Very dangerous removing without to check option existing
    },
    svc_SetRequired: function (lName, isRequre) {
        // Set the attribute required
        var lAttrib = Xrm.Page.getAttribute(lName);

        if (lAttrib) lAttrib.setRequiredLevel(isRequre ? 'required' : 'none');
    },
    svc_SubValue: function (theAttr, valueName, sbvalName, checkBool) {
        if (typeof sbvalName == 'undefined' || sbvalName == null || sbvalName == '') sbvalName = 'value';
        var ilValue = FetchSDK.svc_GetAttr(theAttr, valueName),
            subValue = FetchSDK.svc_GetAttr(ilValue, sbvalName);

        if (subValue == '' && checkBool && FetchSDK.svc_GetAttr(ilValue, 'formattedValue') != '') return false
        else return (subValue == '') ? null : subValue
    },
    svc_ViewReqFld: function (theAttr, flagsSrc, visFlag, reqFlag) {
        var reqClass = Incdnt.svc_SubValue(flagsSrc, reqFlag, '', true),
            visClass = Incdnt.svc_SubValue(flagsSrc, visFlag, '', true);

        Incdnt.svc_CntrlVisible(theAttr, (visClass != null ? visClass : true));
        Incdnt.svc_SetRequired(theAttr, (visClass && reqClass != null ? reqClass : false));
    },
    svc_TakeAttribs: function (theEntity) {
        var atr_SetOf = 'attributes',
            err_Message = FetchSDK.svc_GetAttr(theEntity, 'error'); // ERROR can be

        return (theEntity && err_Message == '' && atr_SetOf in theEntity) ? theEntity[atr_SetOf] : null
    },
    svc_GET_Sbjct: function (sbjctUid) {
        if (typeof sbjctUid == 'undefined' || sbjctUid == null || sbjctUid == '') return;

        return Incdnt.svc_TakeAttribs(FetchSDK.RetrieveSngl(sbjctUid, Incdnt.ent_Subj, Incdnt.atr_SET_Sbjct))
    },
    svc_GET_CustAction: function (actName, reqParams) {
        if (typeof actName == 'undefined' || actName == null || actName == '') actName = 'ing_config_IncidentParam';

        return Incdnt.svc_TakeAttribs(FetchSDK.OrgRequest(actName, reqParams))
    },
    svc_cnfgParam: function (inPriority) {
        if (typeof inPriority == 'undefined' || inPriority == null || isNaN(inPriority)) inPriority = 2;

        return FetchSDK.svc_KeyValue('inPriority', 'c:int', inPriority, 'b')
    },
    OnLoad: function () {
        var flyoutInpt = document.getElementById('ing_fullname');
        if (flyoutInpt) {
            flyoutInpt.setAttribute('compositionlinkcontrolid', 'customerpane_qfc_customerpane_qfc_contact_customerpane_qfc_customerpane_qfc_contact_fullname_compositionLinkControl');
            flyoutInpt.setAttribute('hascompositedata', 'true');
            flyoutInpt.setAttribute('compositerequiredlevel', 'Required');
        }

        Mscrm.CompositeDataControlUtilities.$Q.push('mobilephone');
        Mscrm.CompositeDataControlUtilities.checkForComposeFullName = function Mscrm_CompositeDataControlUtilities$checkForComposeFullName(linkControlObj) {
            linkControlObj.get_flyOutDialog().set_hideFlyOutOnConfirmClick(true);
        }

        Incdnt.MAIN_Config = Incdnt.svc_GET_CustAction('', Incdnt.svc_cnfgParam(Incdnt.svc_AttrValue(Incdnt.atr_Prior)));

        Incdnt.svc_RemOptions(Incdnt.atr_Prior, [3]);
        Incdnt.svc_RemOptions(Incdnt.atr_FirstResp, (Incdnt.svc_AttrValue(Incdnt.atr_FirstResp) == 3 ? [4] : [3, 4]));
        
        if (Incdnt.svc_AttrValue('statuscode') == 1) {
            Incdnt.OnChng_Subject();
            Incdnt.OnChng_Priority(true);
        }
    },
    OnSave: function () {
        if (Xrm.Page.ui.getFormType() == 1) Incdnt.OnChng_Priority(true);
    },
    OnChng_Subject: function () {
        // OnChange of Subject & SubSubject
        var lookSubj = Incdnt.svc_AttribLookValue(Incdnt.atr_Subj_Id),
            lookSubsub = Incdnt.svc_AttribLookValue(Incdnt.atr_Subsub_Id);

        if (lookSubsub == null && lookSubj == null) return;

        var SET_subjAttr = Incdnt.svc_GET_Sbjct((lookSubsub == null) ? lookSubj.id : lookSubsub.id);

        if (!SET_subjAttr) return;

        var caseType = Incdnt.svc_SubValue(SET_subjAttr, Incdnt.atr_SET_Sbjct[0]);
        if (caseType) Incdnt.svc_UpdateValue('casetypecode', caseType, true);

        Incdnt.svc_ViewReqFld('ing_classifier_id', SET_subjAttr, Incdnt.atr_SET_Sbjct[2], Incdnt.atr_SET_Sbjct[1]);
        Incdnt.svc_ViewReqFld('ing_equipment', SET_subjAttr, Incdnt.atr_SET_Sbjct[4], Incdnt.atr_SET_Sbjct[3]);

        var teamUid = Incdnt.svc_SubValue(SET_subjAttr, Incdnt.atr_SET_Sbjct[5], 'id'),
            teamName = Incdnt.svc_SubValue(SET_subjAttr, Incdnt.atr_SET_Sbjct[5], 'name');

        if (!teamUid) teamUid = Incdnt.svc_SubValue(Incdnt.MAIN_Config, 'defaultTeam', 'id');

        Incdnt.svc_UpdateValue('ownerid', [{
            id: (teamUid) ? teamUid : Incdnt.MAIN_defTeam.id,
            name: (teamName) ? teamName : Incdnt.MAIN_defTeam.name,
            entityType: 'team'
        }]);
    },
    OnChng_Priority: function (configUpdated) {
        // On Change of priority for term assignment
        if (Xrm.Page.ui.getFormType() == 0 || Xrm.Page.ui.getFormType() > 2) return;

        if (!configUpdated) Incdnt.MAIN_Config = Incdnt.svc_GET_CustAction('', Incdnt.svc_cnfgParam(Incdnt.svc_AttrValue(Incdnt.atr_Prior)));

        var incdntDate = Incdnt.svc_AttrValue('createdon');

        var clndrArr = FetchSDK.RetrieveMulti(FetchSDK.getPack('calendar', ['calendarid', 'primaryuserid', 'type', 'typename'], ['createdon'],
                                {
                                    conditions: [
                                        { attribute: 'primaryuserid', operator: 'eq', value: '{F6506E64-6133-E811-80EC-005056BA18B6}' },
                                        { attribute: 'type', operator: 'eq', value: '0' }
                                    ]
                                }));

        if (!incdntDate) incdntDate = new Date();
    }
}
