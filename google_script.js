/**
 * Google Apps Script for Approval System Integration
 * 
 * 1. Go to extensions > Apps Script.
 * 2. Paste this code into Code.gs
 * 3. Deploy > New Deployment > Type: Web App
 * 4. Execute as: Me
 * 5. Who has access: Anyone
 * 6. Copy the URL and provide it to the server.
 */

// Configuration: Sheet IDs (Extracted from user provided URLs)
const LOG_SHEET_ID = "18avCkH277cXbus15y0mPqB92WMh0IbVe0nKXxBC6IJE";
const USER_SHEET_ID = "1SL7VYzQet3Ne9omWLK5o-1FfOzaXAwHPSZrZplfFUJY";
const ATTACHMENT_FOLDER_ID = "1BrjhQsV8kr8CGPM28KDkgzJ-IqMmzMCV";
const PRINT_SNAPSHOT_FOLDER_ID = "1umcoZhsP0cd9zk9d7lDJUGIA9DcA5-dM";
const DELETED_ARCHIVE_FOLDER_ID = "1NpYmr1xTdSrappZRC8_UiinoBE73LkTP";
const APPROVAL_STAMP_MAX_WIDTH = 34;   // Narrow approval cells in template
const APPROVAL_STAMP_MAX_HEIGHT = 34;  // Keep square seals fully visible
const SCRIPT_BUILD = "trip-result-copy-2026-03-03-1600";

// ??????λ땾???紐꾩춿疫꿸퀣肉????쎈뻬??뤿연 Drive 亦낅슦釉???諭???뤾쉭????
function testDriveAccess() {
    var folder = DriveApp.getFolderById("1ZwpNkDKZKw2EZ9tqR4RCL4ILiaE6BnaG");
    Logger.log("Drive access check folder: " + folder.getName());
    var testFile = folder.createFile("test_permission.txt", "permission test", "text/plain");
    Logger.log("Drive access check file: " + testFile.getName());
    testFile.setTrashed(true);
    Logger.log("Drive access check cleanup complete");
}

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const action = data.action;

        if (action === "log_login") {
            logLogin(data);
            return success({ message: "Logged" });
        }
        else if (action === "sync_user") {
            var result = syncUser(data.user, data.profile_image, data.approval_stamp_image);
            return success({
                message: "User Synced",
                profile_image_url: result.profile_image_url || null,
                approval_stamp_image_url: result.approval_stamp_image_url || null,
                image_error: result.image_error || null,
                approval_stamp_error: result.approval_stamp_error || null
            });
        }
        else if (action === "delete_user") {
            deleteUser(data.username);
            return success({ message: "User Deleted" });
        }
        else if (action === "update_approval_doc") {
            var updateResult = updateApprovalDoc(data);
            return success({ message: "Approval Doc Updated", updated_slots: updateResult.updated_slots || [], errors: updateResult.errors || [] });
        }
        else if (action === "validate_approval_template") {
            var validation = validateApprovalTemplate(data);
            return success(validation);
        }
        else if (action === "populate_draft_doc_fields") {
            var fillResult = populateDraftDocFields(data);
            return success({ message: "Draft Doc Fields Populated", result: fillResult });
        }
        else if (action === "upload_attachments") {
            var uploadResult = uploadAttachments(data);
            return success({ message: "Attachments Uploaded", attachments: uploadResult });
        }
        else if (action === "reset_approval_doc") {
            var resetResult = resetApprovalDoc(data);
            return success({ message: "Approval Doc Reset", reset: resetResult });
        }
        else if (action === "copy_template_doc") {
            var copyResult = copyTemplateDoc(data);
            return success({ message: "Template Copied", copy: copyResult });
        }
        else if (action === "delete_drive_file") {
            var deleteResult = deleteDriveFile(data);
            return success({ message: "Drive File Deleted", drive_delete: deleteResult });
        }
        else if (action === "move_drive_file") {
            var moveResult = moveDriveFile(data);
            return success({ message: "Drive File Moved", drive_move: moveResult });
        }
        else if (action === "create_pdf_snapshot") {
            var snapshotResult = createPdfSnapshot(data);
            return success({ message: "PDF Snapshot Created", snapshot: snapshotResult });
        }
        else if (action === "read_drive_file_base64") {
            var fileReadResult = readDriveFileBase64(data);
            return success({ message: "Drive File Read", file: fileReadResult });
        }

        return error("Unknown action");
    } catch (err) {
        return error(err.message);
    }
}

function doGet(e) {
    // Can be used to fetch users
    const action = e.parameter.action;
    if (action === "get_users") {
        return getUsers();
    }
    return success({ message: "Approval System API" });
}

function logLogin(data) {
    const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    const sheet = ss.getSheets()[0]; // Assume first sheet
    // Format: Timestamp, Username, IP, Status
    sheet.appendRow([new Date(), data.username, data.ip || "Unknown", "Success"]);
}

function syncUser(user, profileImage, approvalStampImage) {
    const ss = SpreadsheetApp.openById(USER_SHEET_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    var result = { profile_image_url: null, approval_stamp_image_url: null };

    // Upload profile image to Drive if provided
    var imageUrl = null;
    if (profileImage && profileImage.data) {
        var uploadResult = uploadImageToDrive(profileImage, user.username, "profile");
        if (uploadResult.url) {
            imageUrl = uploadResult.url;
            result.profile_image_url = imageUrl;
        } else {
            result.image_error = uploadResult.error || "Unknown upload error";
        }
    }

    // Upload approval stamp to Drive if provided
    var stampImageUrl = null;
    if (approvalStampImage && approvalStampImage.data) {
        var stampUploadResult = uploadImageToDrive(approvalStampImage, user.username, "stamp");
        if (stampUploadResult.url) {
            stampImageUrl = stampUploadResult.url;
            result.approval_stamp_image_url = stampImageUrl;
        } else {
            result.approval_stamp_error = stampUploadResult.error || "Unknown upload error";
        }
    }

    // Headers: Username, Password, FullName, Department, JobTitle, Role, TotalLeave, UsedLeave, ProfileImageURL, ApprovalStampImageURL
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] == user.username) {
            rowIndex = i + 1;
            break;
        }
    }

    // If no new image uploaded, preserve existing URL from sheet
    var existingImageUrl = (rowIndex > 0 && data[rowIndex - 1].length >= 9) ? data[rowIndex - 1][8] : "";
    var existingStampImageUrl = (rowIndex > 0 && data[rowIndex - 1].length >= 10) ? data[rowIndex - 1][9] : "";
    var finalImageUrl = imageUrl || user.profile_image_url || existingImageUrl || "";
    var finalStampImageUrl = stampImageUrl || user.approval_stamp_image_url || existingStampImageUrl || "";
    if (!result.profile_image_url && finalImageUrl) {
        result.profile_image_url = finalImageUrl;
    }
    if (!result.approval_stamp_image_url && finalStampImageUrl) {
        result.approval_stamp_image_url = finalStampImageUrl;
    }

    const rowData = [
        user.username,
        "",
        user.full_name,
        user.department,
        user.job_title || "",
        user.role,
        user.total_leave || 0,
        user.used_leave || 0,
        finalImageUrl,
        finalStampImageUrl
    ];

    if (rowIndex > 0) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
        sheet.appendRow(rowData);
    }

    return result;
}

function uploadAttachments(data) {
    var files = Array.isArray(data.files) ? data.files : [];
    if (!files.length) return [];

    var folder = DriveApp.getFolderById(ATTACHMENT_FOLDER_ID);
    var out = [];

    for (var i = 0; i < files.length; i++) {
        var item = files[i] || {};
        if (!item.data) continue;

        var name = String(item.name || ("attachment_" + (i + 1))).trim();
        var mime = String(item.type || item.mime_type || "application/octet-stream").trim() || "application/octet-stream";
        var decoded = Utilities.base64Decode(item.data);
        var blob = Utilities.newBlob(decoded, mime, name);

        var file = folder.createFile(blob);
        try {
            file.setName(name);
        } catch (e) {}

        out.push({
            file_id: file.getId(),
            id: file.getId(),
            name: file.getName(),
            mime_type: file.getMimeType(),
            size: file.getSize(),
            web_view_url: file.getUrl()
        });
    }

    return out;
}

function copyTemplateDoc(data) {
    var templateDocId = String(data.template_doc_id || data.templateDocId || "").trim();
    var targetFolderId = String(data.target_folder_id || data.targetFolderId || "").trim();
    var title = String(data.title || "").trim();
    if (!templateDocId) throw new Error("template_doc_id is required");
    if (!targetFolderId) throw new Error("target_folder_id is required");

    var templateFile = DriveApp.getFileById(templateDocId);
    var targetFolder = DriveApp.getFolderById(targetFolderId);
    var copyName = title || (templateFile.getName() + "_copy");
    var copied = templateFile.makeCopy(copyName, targetFolder);
    var copiedId = copied.getId();
    return {
        id: copiedId,
        doc_id: copiedId,
        file_id: copiedId,
        name: copied.getName(),
        web_view_url: copied.getUrl(),
        edit_url: "https://docs.google.com/document/d/" + copiedId + "/edit"
    };
}

function validateApprovalTemplate(data) {
    var docId = String(data.doc_id || data.docId || "").trim();
    if (!docId) throw new Error("doc_id is required");
    var requiredSlots = Number(data.required_slots || 1);
    if (!requiredSlots || requiredSlots < 1) requiredSlots = 1;

    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    var missing = [];
    var found = [];

    for (var i = 1; i <= requiredSlots; i++) {
        var nameToken = "{{APPR" + i + "_NAME}}";
        var stampToken = "{{APPR" + i + "_STAMP}}";
        if (body.findText(escapeRegExp(nameToken))) found.push(nameToken);
        else missing.push(nameToken);
        if (body.findText(escapeRegExp(stampToken))) found.push(stampToken);
        else missing.push(stampToken);
    }

    return {
        valid: missing.length === 0,
        required_slots: requiredSlots,
        found_tokens: found,
        missing_tokens: missing
    };
}

function populateDraftDocFields(data) {
    var docId = String(data.doc_id || data.docId || "").trim();
    if (!docId) throw new Error("doc_id is required");

    var title = String(data.title || "").trim();
    var issueDate = String(data.issue_date || data.issueDate || "").trim();
    var issueCode = String(data.issue_code || data.issueCode || "").trim();
    var docFormType = String(data.doc_form_type || data.docFormType || "").trim();
    var recipientText = String(data.recipient_text || data.recipientText || "").trim();
    var leaveDept = String(data.leave_dept || data.leaveDept || "").trim();
    var leaveName = String(data.leave_name || data.leaveName || "").trim();
    var leaveJobTitle = String(data.leave_job_title || data.leaveJobTitle || "").trim();
    var leaveType = String(data.leave_type || data.leaveType || "").trim();
    var leavePeriod = String(data.leave_period || data.leavePeriod || "").trim();
    var leaveTotalDays = String(data.leave_total_days || data.leaveTotalDays || "").trim();
    var leaveUsedDays = String(data.leave_used_days || data.leaveUsedDays || "").trim();
    var leaveRemainDays = String(data.leave_remain_days || data.leaveRemainDays || "").trim();
    var leaveReason = String(data.leave_reason || data.leaveReason || "").trim();
    var leaveSubstituteName = String(data.leave_substitute_name || data.leaveSubstituteName || "").trim();
    var leaveSubstituteWork = String(data.leave_substitute_work || data.leaveSubstituteWork || "").trim();
    var draftDate = String(data.draft_date || data.draftDate || "").trim();
    var overtimeDept = String(data.overtime_dept || data.overtimeDept || "").trim();
    var overtimeName = String(data.overtime_name || data.overtimeName || "").trim();
    var overtimeJobTitle = String(data.overtime_job_title || data.overtimeJobTitle || "").trim();
    var overtimeType = String(data.overtime_type || data.overtimeType || "").trim();
    var overtimeTime = String(data.overtime_time || data.overtimeTime || "").trim();
    var overtimeHours = String(data.overtime_hours || data.overtimeHours || "").trim();
    var overtimeContent = String(data.overtime_content || data.overtimeContent || "").trim();
    var overtimeEtc = String(data.overtime_etc || data.overtimeEtc || "").trim();
    var tripDepartment = String(data.trip_department || data.tripDepartment || "").trim();
    var tripJobTitle = String(data.trip_job_title || data.tripJobTitle || "").trim();
    var tripName = String(data.trip_name || data.tripName || "").trim();
    var tripType = String(data.trip_type || data.tripType || "").trim();
    var tripDestination = String(data.trip_destination || data.tripDestination || "").trim();
    var tripPeriod = String(data.trip_period || data.tripPeriod || "").trim();
    var tripTransportation = String(data.trip_transportation || data.tripTransportation || "").trim();
    var tripExpense = String(data.trip_expense || data.tripExpense || "").trim();
    var tripPurpose = String(data.trip_purpose || data.tripPurpose || "").trim();
    var tripResult = String(data.trip_result || data.tripResult || "").trim();
    var educationDepartment = String(data.education_department || data.educationDepartment || "").trim();
    var educationJobTitle = String(data.education_job_title || data.educationJobTitle || "").trim();
    var educationName = String(data.education_name || data.educationName || "").trim();
    var educationTitle = String(data.education_title || data.educationTitle || "").trim();
    var educationCategory = String(data.education_category || data.educationCategory || "").trim();
    var educationProvider = String(data.education_provider || data.educationProvider || "").trim();
    var educationLocation = String(data.education_location || data.educationLocation || "").trim();
    var educationPeriod = String(data.education_period || data.educationPeriod || "").trim();
    var educationPurpose = String(data.education_purpose || data.educationPurpose || "").trim();
    var educationTuitionDetail = String(data.education_tuition_detail || data.educationTuitionDetail || "").trim();
    var educationTuitionAmount = String(data.education_tuition_amount || data.educationTuitionAmount || "").trim();
    var educationMaterialDetail = String(data.education_material_detail || data.educationMaterialDetail || "").trim();
    var educationMaterialAmount = String(data.education_material_amount || data.educationMaterialAmount || "").trim();
    var educationTransportDetail = String(data.education_transport_detail || data.educationTransportDetail || "").trim();
    var educationTransportAmount = String(data.education_transport_amount || data.educationTransportAmount || "").trim();
    var educationOtherDetail = String(data.education_other_detail || data.educationOtherDetail || "").trim();
    var educationOtherAmount = String(data.education_other_amount || data.educationOtherAmount || "").trim();
    var educationBudgetSubject = String(data.education_budget_subject || data.educationBudgetSubject || "").trim();
    var educationFundingSource = String(data.education_funding_source || data.educationFundingSource || "").trim();
    var educationPaymentMethod = String(data.education_payment_method || data.educationPaymentMethod || "").trim();
    var educationSupportBudget = String(data.education_support_budget || data.educationSupportBudget || "").trim();
    var educationUsedBudget = String(data.education_used_budget || data.educationUsedBudget || "").trim();
    var educationRemainBudget = String(data.education_remain_budget || data.educationRemainBudget || "").trim();
    var educationCompanion = String(data.education_companion || data.educationCompanion || "").trim();
    var educationOrdered = String(data.education_ordered || data.educationOrdered || "").trim();
    var educationSuggestion = String(data.education_suggestion || data.educationSuggestion || "").trim();
    var educationContent = String(data.education_content || data.educationContent || "").trim();
    var educationApplyPoint = String(data.education_apply_point || data.educationApplyPoint || "").trim();

    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();

    var result = {
        title: title ? writeDraftTitle(body, title) : { updated: false, reason: "empty" },
        issue_date: issueDate ? writeDraftFieldByLabelOrToken(body, {
            token: "{{ISSUE_DATE}}",
            labelCellTexts: ["\uC2DC\uD589\uC77C"],
            value: issueDate
        }) : { updated: false, reason: "empty" },
        issue_code: issueCode ? writeDraftFieldByLabelOrToken(body, {
            token: "{{ISSUE_NO}}",
            labelCellTexts: ["\uC2DC\uD589"],
            value: issueCode
        }) : { updated: false, reason: "empty" },
        doc_form_type: docFormType ? writeDraftFieldByLabelOrToken(body, {
            token: "{{DOC_FORM_TYPE}}",
            labelCellTexts: ["\uBB38\uC11C\uC591\uC2DD", "\uBB38\uC11C\uC720\uD615"],
            value: docFormType
        }) : { updated: false, reason: "empty" },
        issue_receiver: recipientText ? writeDraftFieldByLabelOrToken(body, {
            token: "{{ISSUE_RECEIVER}}",
            labelCellTexts: ["\uC2DC\uD589\uBB38\uC218\uC2E0", "\uC218\uC2E0"],
            value: recipientText
        }) : { updated: false, reason: "empty" },
        draft_date: (draftDate || issueDate) ? writeDraftFieldByLabelOrToken(body, {
            token: "{{DRAFT_DATE}}",
            labelCellTexts: ["\uAE30\uC548\uC77C\uC790", "\uAE30\uC548\uC77C"],
            value: draftDate || issueDate
        }) : { updated: false, reason: "empty" },
        leave_dept: leaveDept ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_DEPT}}",
            labelCellTexts: ["\uBD80\uC11C"],
            value: leaveDept
        }) : { updated: false, reason: "empty" },
        leave_name: leaveName ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_NAME}}",
            labelCellTexts: ["\uC131\uBA85"],
            value: leaveName
        }) : { updated: false, reason: "empty" },
        leave_job_title: leaveJobTitle ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_JOB_TITLE}}",
            labelCellTexts: ["\uC9C1\uC704", "\uC9C1\uAE09"],
            value: leaveJobTitle
        }) : { updated: false, reason: "empty" },
        leave_type: leaveType ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_TYPE}}",
            labelCellTexts: ["\uD734\uAC00\uD615\uD0DC"],
            value: leaveType
        }) : { updated: false, reason: "empty" },
        leave_period: leavePeriod ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_PERIOD}}",
            labelCellTexts: ["\uD734\uAC00\uAE30\uAC04"],
            value: leavePeriod
        }) : { updated: false, reason: "empty" },
        leave_total_days: leaveTotalDays ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_TOTAL_DAYS}}",
            labelCellTexts: ["\uC5F0\uCC28\uBD80\uC5EC\uC77C\uC218"],
            value: leaveTotalDays
        }) : { updated: false, reason: "empty" },
        leave_used_days: leaveUsedDays ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_USED_DAYS}}",
            labelCellTexts: ["\uC5F0\uCC28\uC0AC\uC6A9\uC77C\uC218"],
            value: leaveUsedDays
        }) : { updated: false, reason: "empty" },
        leave_remain_days: leaveRemainDays ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_REMAIN_DAYS}}",
            labelCellTexts: ["\uC5F0\uCC28\uC794\uC5EC\uC77C\uC218"],
            value: leaveRemainDays
        }) : { updated: false, reason: "empty" },
        leave_reason: leaveReason ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_REASON}}",
            labelCellTexts: ["\uD734\uAC00\uC0AC\uC720"],
            value: leaveReason
        }) : { updated: false, reason: "empty" },
        leave_substitute_name: leaveSubstituteName ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_SUBSTITUTE_NAME}}",
            labelCellTexts: ["\uB300\uC9C1\uC790"],
            value: leaveSubstituteName
        }) : { updated: false, reason: "empty" },
        leave_substitute_work: leaveSubstituteWork ? writeDraftFieldByLabelOrToken(body, {
            token: "{{LEAVE_SUBSTITUTE_WORK}}",
            labelCellTexts: ["\uB300\uC9C1\uC790\uC5C5\uBB34\uB0B4\uC6A9", "\uB300\uC9C1\uC790 \uC5C5\uBB34\uB0B4\uC6A9"],
            value: leaveSubstituteWork
        }) : { updated: false, reason: "empty" },
        overtime_dept: overtimeDept ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_DEPT}}",
            labelCellTexts: ["\uC18C\uC18D", "\uBD80\uC11C"],
            value: overtimeDept
        }) : { updated: false, reason: "empty" },
        overtime_name: overtimeName ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_NAME}}",
            labelCellTexts: ["\uC131\uBA85"],
            value: overtimeName
        }) : { updated: false, reason: "empty" },
        overtime_job_title: overtimeJobTitle ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_JOB_TITLE}}",
            labelCellTexts: ["\uC9C1\uC704", "\uC9C1\uAE09"],
            value: overtimeJobTitle
        }) : { updated: false, reason: "empty" },
        overtime_type: overtimeType ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_TYPE}}",
            labelCellTexts: ["\uD615\uD0DC"],
            value: overtimeType
        }) : { updated: false, reason: "empty" },
        overtime_time: overtimeTime ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_TIME}}",
            labelCellTexts: ["\uC5F0\uC7A5\uADFC\uB85C \uC2DC\uAC04", "\uC5F0\uC7A5\uADFC\uB85C\uC2DC\uAC04"],
            value: overtimeTime
        }) : { updated: false, reason: "empty" },
        overtime_hours: overtimeHours ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_HOURS}}",
            labelCellTexts: [],
            value: overtimeHours
        }) : { updated: false, reason: "empty" },
        overtime_content: overtimeContent ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_CONTENT}}",
            labelCellTexts: ["\uC5F0\uC7A5\uADFC\uB85C \uB0B4\uC6A9", "\uC5F0\uC7A5\uADFC\uB85C\uB0B4\uC6A9"],
            value: overtimeContent
        }) : { updated: false, reason: "empty" },
        overtime_etc: overtimeEtc ? writeDraftFieldByLabelOrToken(body, {
            token: "{{OT_ETC}}",
            labelCellTexts: ["\uAE30\uD0C0"],
            value: overtimeEtc
        }) : { updated: false, reason: "empty" },
        trip_department: tripDepartment ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_DEPT}}",
            labelCellTexts: ["\uC18C\uC18D", "\uBD80\uC11C"],
            value: tripDepartment
        }) : { updated: false, reason: "empty" },
        trip_job_title: tripJobTitle ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_JOB_TITLE}}",
            labelCellTexts: ["\uC9C1\uC704", "\uC9C1\uAE09"],
            value: tripJobTitle
        }) : { updated: false, reason: "empty" },
        trip_name: tripName ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_NAME}}",
            labelCellTexts: ["\uC131\uBA85"],
            value: tripName
        }) : { updated: false, reason: "empty" },
        trip_type: tripType ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_TYPE}}",
            labelCellTexts: ["\uCD9C\uC7A5\uC885\uB958"],
            value: tripType
        }) : { updated: false, reason: "empty" },
        trip_destination: tripDestination ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_DESTINATION}}",
            labelCellTexts: ["\uCD9C\uC7A5\uC9C0"],
            value: tripDestination
        }) : { updated: false, reason: "empty" },
        trip_period: tripPeriod ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_PERIOD}}",
            labelCellTexts: ["\uCD9C\uC7A5\uAE30\uAC04"],
            value: tripPeriod
        }) : { updated: false, reason: "empty" },
        trip_transportation: tripTransportation ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_TRANSPORT}}",
            labelCellTexts: ["\uAD50\uD1B5\uC218\uB2E8"],
            value: tripTransportation
        }) : { updated: false, reason: "empty" },
        trip_expense: tripExpense ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_EXPENSE}}",
            labelCellTexts: ["\uCD9C\uC7A5\uBE44"],
            value: tripExpense
        }) : { updated: false, reason: "empty" },
        trip_purpose: tripPurpose ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_PURPOSE}}",
            labelCellTexts: ["\uCD9C\uC7A5\uBAA9\uC801"],
            value: tripPurpose
        }) : { updated: false, reason: "empty" },
        trip_result: tripResult ? writeDraftFieldByLabelOrToken(body, {
            token: "{{TRIP_RESULT}}",
            labelCellTexts: ["\uCD9C\uC7A5\uACB0\uACFC"],
            value: tripResult
        }) : { updated: false, reason: "empty" },
        education_department: educationDepartment ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_DEPT}}",
            labelCellTexts: ["\uC18C\uC18D", "\uBD80\uC11C"],
            value: educationDepartment
        }) : { updated: false, reason: "empty" },
        education_job_title: educationJobTitle ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_JOB_TITLE}}",
            labelCellTexts: ["\uC9C1\uC704", "\uC9C1\uAE09"],
            value: educationJobTitle
        }) : { updated: false, reason: "empty" },
        education_name: educationName ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_NAME}}",
            labelCellTexts: ["\uC131\uBA85"],
            value: educationName
        }) : { updated: false, reason: "empty" },
        education_title: educationTitle ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_TITLE}}",
            labelCellTexts: ["\uAD50\uC721\uBA85"],
            value: educationTitle
        }) : { updated: false, reason: "empty" },
        education_category: educationCategory ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_CATEGORY}}",
            labelCellTexts: ["\uAD50\uC721\uBD84\uB958"],
            value: educationCategory
        }) : { updated: false, reason: "empty" },
        education_provider: educationProvider ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_PROVIDER}}",
            labelCellTexts: ["\uAD50\uC721\uAE30\uAD00"],
            value: educationProvider
        }) : { updated: false, reason: "empty" },
        education_location: educationLocation ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_LOCATION}}",
            labelCellTexts: ["\uAD50\uC721\uC7A5\uC18C"],
            value: educationLocation
        }) : { updated: false, reason: "empty" },
        education_period: educationPeriod ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_PERIOD}}",
            labelCellTexts: ["\uAD50\uC721\uC77C\uC2DC", "\uAD50\uC721\uAE30\uAC04"],
            value: educationPeriod
        }) : { updated: false, reason: "empty" },
        education_purpose: educationPurpose ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_PURPOSE}}",
            labelCellTexts: ["\uAD50\uC721\uBAA9\uC801"],
            value: educationPurpose
        }) : { updated: false, reason: "empty" },
        education_tuition_detail: educationTuitionDetail ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_TUITION_DETAIL}}",
            labelCellTexts: [],
            value: educationTuitionDetail
        }) : { updated: false, reason: "empty" },
        education_tuition_amount: educationTuitionAmount ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_TUITION_AMOUNT}}",
            labelCellTexts: [],
            value: educationTuitionAmount
        }) : { updated: false, reason: "empty" },
        education_material_detail: educationMaterialDetail ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_MATERIAL_DETAIL}}",
            labelCellTexts: [],
            value: educationMaterialDetail
        }) : { updated: false, reason: "empty" },
        education_material_amount: educationMaterialAmount ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_MATERIAL_AMOUNT}}",
            labelCellTexts: [],
            value: educationMaterialAmount
        }) : { updated: false, reason: "empty" },
        education_transport_detail: educationTransportDetail ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_TRANSPORT_DETAIL}}",
            labelCellTexts: [],
            value: educationTransportDetail
        }) : { updated: false, reason: "empty" },
        education_transport_amount: educationTransportAmount ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_TRANSPORT_AMOUNT}}",
            labelCellTexts: [],
            value: educationTransportAmount
        }) : { updated: false, reason: "empty" },
        education_other_detail: educationOtherDetail ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_OTHER_DETAIL}}",
            labelCellTexts: [],
            value: educationOtherDetail
        }) : { updated: false, reason: "empty" },
        education_other_amount: educationOtherAmount ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_OTHER_AMOUNT}}",
            labelCellTexts: [],
            value: educationOtherAmount
        }) : { updated: false, reason: "empty" },
        education_budget_subject: educationBudgetSubject ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_BUDGET_SUBJECT}}",
            labelCellTexts: ["\uC608\uC0B0\uACFC\uBAA9"],
            value: educationBudgetSubject
        }) : { updated: false, reason: "empty" },
        education_funding_source: educationFundingSource ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_FUNDING_SOURCE}}",
            labelCellTexts: ["\uC790\uAE08\uC6D0\uCC9C"],
            value: educationFundingSource
        }) : { updated: false, reason: "empty" },
        education_payment_method: educationPaymentMethod ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_PAYMENT_METHOD}}",
            labelCellTexts: ["\uACB0\uC7AC\uBC29\uBC95"],
            value: educationPaymentMethod
        }) : { updated: false, reason: "empty" },
        education_support_budget: educationSupportBudget ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_SUPPORT_BUDGET}}",
            labelCellTexts: ["\uC9C0\uC6D0\uC608\uC0B0(\uC6D0)", "\uC9C0\uC6D0\uC608\uC0B0"],
            value: educationSupportBudget
        }) : { updated: false, reason: "empty" },
        education_used_budget: educationUsedBudget ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_USED_BUDGET}}",
            labelCellTexts: ["\uC0AC\uC6A9\uC608\uC0B0(\uC6D0)", "\uC0AC\uC6A9\uC608\uC0B0"],
            value: educationUsedBudget
        }) : { updated: false, reason: "empty" },
        education_remain_budget: educationRemainBudget ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_REMAIN_BUDGET}}",
            labelCellTexts: ["\uC794\uC5EC\uAE08\uC561(\uC6D0)", "\uC794\uC5EC\uAE08\uC561"],
            value: educationRemainBudget
        }) : { updated: false, reason: "empty" },
        education_companion: educationCompanion ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_COMPANION}}",
            labelCellTexts: ["\uB3D9\uD589\uC790"],
            value: educationCompanion
        }) : { updated: false, reason: "empty" },
        education_ordered: educationOrdered ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_ORDERED}}",
            labelCellTexts: ["\uAE30\uAD00 \uBA85\uB839 \uC5EC\uBD80", "\uAE30\uAD00\uBA85\uB839\uC5EC\uBD80"],
            value: educationOrdered
        }) : { updated: false, reason: "empty" },
        education_suggestion: educationSuggestion ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_SUGGESTION}}",
            labelCellTexts: ["\uAC74\uC758\uC0AC\uD56D"],
            value: educationSuggestion
        }) : { updated: false, reason: "empty" },
        education_content: educationContent ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_CONTENT}}",
            labelCellTexts: ["\uAD50\uC721\uB0B4\uC6A9"],
            value: educationContent
        }) : { updated: false, reason: "empty" },
        education_apply_point: educationApplyPoint ? writeDraftFieldByLabelOrToken(body, {
            token: "{{EDU_APPLY_POINT}}",
            labelCellTexts: ["\uC801\uC6A9\uC810"],
            value: educationApplyPoint
        }) : { updated: false, reason: "empty" }
    };

    doc.saveAndClose();
    return result;
}

function updateApprovalDoc(data) {
    var docId = (data.doc_id || data.docId || "").toString().trim();
    if (!docId) throw new Error("doc_id is required");

    var slots = Array.isArray(data.slots) ? data.slots : [];
    var totalSlots = Number(data.total_slots || data.totalSlots || 4);
    if (!totalSlots || totalSlots < 1) totalSlots = 4;
    var usedSlots = Number(data.used_slots || data.usedSlots || slots.length || 0);
    if (!usedSlots || usedSlots < 0) usedSlots = slots.length || 0;
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    var result = { updated_slots: [], errors: [] };
    var effectiveTotalSlots = totalSlots;

    function findWorkingTotalSlots(preferred) {
        var candidates = [];
        if (preferred && preferred > 0) candidates.push(preferred);
        for (var c = 1; c <= 6; c++) {
            if (candidates.indexOf(c) < 0) candidates.push(c);
        }
        for (var i = 0; i < candidates.length; i++) {
            var n = candidates[i];
            try {
                if (locateApprovalTableLayout(body, n)) return n;
            } catch (e) {}
        }
        return preferred;
    }
    effectiveTotalSlots = findWorkingTotalSlots(totalSlots);

    try {
        result.cleaned_misplaced_rows = cleanupMisplacedApprovalRows(body, Math.max(effectiveTotalSlots, 4));
    } catch (e) {
        result.errors.push({ slot_index: 0, message: "misplaced approval cleanup failed: " + e.message });
    }
    // Re-evaluate after cleanup because the table layout may become detectable only after removing polluted rows.
    effectiveTotalSlots = findWorkingTotalSlots(effectiveTotalSlots);
    result.effective_total_slots = effectiveTotalSlots;

    for (var i = 0; i < slots.length; i++) {
        var slot = slots[i] || {};
        slot.total_slots = effectiveTotalSlots;
        try {
            var applied = applyApprovalSlotToDocument(body, slot);
            result.updated_slots.push(applied || Number(slot.slot_index || slot.slot || slot.index || 0));
        } catch (e) {
            result.errors.push({
                slot_index: Number(slot.slot_index || slot.slot || slot.index || 0),
                message: e.message
            });
        }
    }

    // Clear slots that are not part of the current approval line (e.g. APPR3/APPR4 for a 2-person line).
    // Keep placeholders only for active slots so future approvals can still stamp correctly.
    try {
        var cleared = clearUnusedApprovalSlots(body, effectiveTotalSlots, usedSlots);
        result.cleared_unused_slots = cleared;
    } catch (e) {
        result.errors.push({ slot_index: 0, message: "unused slot clear failed: " + e.message });
    }

    doc.saveAndClose();
    return result;
}

function resetApprovalDoc(data) {
    var docId = String(data.doc_id || data.docId || "").trim();
    if (!docId) throw new Error("doc_id is required");
    var totalSlots = Number(data.total_slots || 4);
    if (!totalSlots || totalSlots < 1) totalSlots = 4;

    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    var cleanupRows = [];
    try {
        cleanupRows = cleanupMisplacedApprovalRows(body, Math.max(totalSlots, 4));
    } catch (e) {
        // best effort
    }
    var resetCount = 0;
    for (var i = 1; i <= totalSlots; i++) {
        resetCount += resetApprovalSlotByTokens(body, i) ? 1 : 0;
    }

    // Fallback for strictly tabular templates if token-based search failed unexpectedly.
    var layout = null;
    var effectiveTotalSlots = totalSlots;
    if (resetCount === 0) {
        var candidates = [];
        candidates.push(totalSlots);
        for (var c = 1; c <= 6; c++) {
            if (candidates.indexOf(c) < 0) candidates.push(c);
        }
        for (var idx = 0; idx < candidates.length; idx++) {
            var slotsTry = candidates[idx];
            layout = locateApprovalTableLayout(body, slotsTry);
            if (layout) {
                effectiveTotalSlots = slotsTry;
                break;
            }
        }
        if (!layout) {
            for (var ex = 0; ex < candidates.length; ex++) {
                var exSlotsTry = candidates[ex];
                layout = locateApprovalTableLayoutInExcludedTables(body, exSlotsTry);
                if (layout) {
                    effectiveTotalSlots = exSlotsTry;
                    break;
                }
            }
        }
        if (!layout) {
            layout = createFallbackApprovalTableLayout(body, totalSlots);
            effectiveTotalSlots = totalSlots;
        }
        for (var j = 1; j <= effectiveTotalSlots; j++) {
            var nameCell = layout.table.getRow(layout.nameRowIndex).getCell(layout.baseColIndex + (j - 1));
            var stampCell = layout.table.getRow(layout.stampRowIndex).getCell(layout.baseColIndex + (j - 1));
            setCellText(nameCell, "{{APPR" + j + "_NAME}}");
            setCellText(stampCell, "{{APPR" + j + "_STAMP}}");
        }
        resetCount = effectiveTotalSlots;
    }

    doc.saveAndClose();
    return {
        doc_id: docId,
        total_slots: totalSlots,
        effective_total_slots: effectiveTotalSlots,
        cleaned_misplaced_rows: cleanupRows,
        reset_slots: resetCount,
        name_row_index: layout ? layout.nameRowIndex : null,
        stamp_row_index: layout ? layout.stampRowIndex : null,
        base_col_index: layout ? layout.baseColIndex : null
    };
}

function clearUnusedApprovalSlots(body, totalSlots, usedSlots) {
    var cleared = [];
    if (!totalSlots || totalSlots < 1) return cleared;
    if (!usedSlots || usedSlots < 0) usedSlots = 0;

    var layout = locateApprovalTableLayout(body, totalSlots);
    for (var i = usedSlots + 1; i <= totalSlots; i++) {
        var nameCleared = false;
        var stampCleared = false;

        var nameToken = "{{APPR" + i + "_NAME}}";
        var nameRange = body.findText(escapeRegExp(nameToken));
        if (nameRange) {
            var nameCell = findParentTableCell(nameRange.getElement().asText());
            if (nameCell) {
                setCellText(nameCell, "");
                nameCleared = true;
            }
        }
        if (!nameCleared && layout) {
            var layoutNameCell = getNameCellFromLayout(layout, i);
            if (layoutNameCell) {
                setCellText(layoutNameCell, "");
                nameCleared = true;
            }
        }

        var stampToken = "{{APPR" + i + "_STAMP}}";
        var stampRange = body.findText(escapeRegExp(stampToken));
        if (stampRange) {
            var stampCell = findParentTableCell(stampRange.getElement().asText());
            if (stampCell) {
                setCellText(stampCell, "");
                stampCleared = true;
            }
        }
        if (!stampCleared && layout) {
            var layoutStampCell = getStampCellFromLayout(layout, i);
            if (layoutStampCell) {
                setCellText(layoutStampCell, "");
                stampCleared = true;
            }
        }

        if (nameCleared || stampCleared) {
            cleared.push(i);
        }
    }
    return cleared;
}

function writeDraftTitle(body, value) {
    var tokenResult = replaceFirstTextCandidate(body, ["{{DOC_TITLE}}"], value);
    if (tokenResult && tokenResult.replaced) {
        return { updated: true, method: "token" };
    }

    var titleCellResult = writeDraftFieldByLabelOrToken(body, {
        token: "{{TITLE}}",
        labelCellTexts: ["\uC81C\uBAA9", "\uC81C \uBAA9"],
        value: value,
        allowParagraphFallback: false
    });
    if (titleCellResult && titleCellResult.updated) {
        return titleCellResult;
    }

    var paraResult = writeValueInParagraphAfterLabel(body, [
        "\uC81C\uBAA9",
        "\uC81C \uBAA9"
    ], value);
    if (paraResult && paraResult.updated) {
        return paraResult;
    }

    return { updated: false, reason: "title target not found" };
}

function writeDraftFieldByLabelOrToken(body, opts) {
    var token = String(opts.token || "").trim();
    var value = String(opts.value || "").trim();
    var labelCellTexts = Array.isArray(opts.labelCellTexts) ? opts.labelCellTexts : [];
    var allowParagraphFallback = opts.allowParagraphFallback !== false;

    if (!value) return { updated: false, reason: "empty" };

    if (token) {
        var tokenReplace = replaceAllTextCandidates(body, [token], value);
        if (tokenReplace && tokenReplace.replaced) {
            return { updated: true, method: "token", count: tokenReplace.count || 1 };
        }
    }

    var adjacent = findAdjacentCellByLabels(body, labelCellTexts);
    if (adjacent) {
        setCellText(adjacent, value);
        return { updated: true, method: "table_label" };
    }

    if (allowParagraphFallback && labelCellTexts.length) {
        var para = writeValueInParagraphAfterLabel(body, labelCellTexts, value);
        if (para && para.updated) return para;
    }

    return { updated: false, reason: "target not found" };
}

function replaceAllTextCandidates(body, candidates, replacement) {
    var count = 0;
    var lastCell = null;
    for (var i = 0; i < candidates.length; i++) {
        var candidate = candidates[i];
        if (!candidate) continue;
        while (true) {
            var range = body.findText(escapeRegExp(candidate));
            if (!range) break;
            var textEl = range.getElement().asText();
            var start = range.getStartOffset();
            var end = range.getEndOffsetInclusive();
            textEl.deleteText(start, end);
            textEl.insertText(start, replacement);
            count++;
            lastCell = findParentTableCell(textEl);
        }
    }
    return { replaced: count > 0, count: count, cell: lastCell };
}

function writeValueInParagraphAfterLabel(body, labels, value) {
    var normalizedLabels = (labels || []).map(function (x) { return normalizeCellText(x); });
    for (var i = 0; i < body.getNumChildren(); i++) {
        var child = body.getChild(i);
        if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
        var p = child.asParagraph();
        var raw = String(p.getText() || "");
        var norm = normalizeCellText(raw);
        for (var j = 0; j < normalizedLabels.length; j++) {
            var label = normalizedLabels[j];
            if (!label) continue;
            if (norm.indexOf(label) === 0) {
                var outLabel = extractVisibleLabel(raw);
                p.setText((outLabel || raw).replace(/\s*$/, "") + " " + value);
                return { updated: true, method: "paragraph_label" };
            }
        }
    }
    return { updated: false, reason: "paragraph label not found" };
}

function extractVisibleLabel(text) {
    var raw = String(text || "");
    var idx = raw.indexOf(":");
    if (idx >= 0) return raw.substring(0, idx + 1);
    return raw;
}

function findAdjacentCellByLabels(body, labels) {
    var normalizedTargets = (labels || []).map(function (x) { return normalizeCellText(x); }).filter(Boolean);
    if (!normalizedTargets.length) return null;

    for (var i = 0; i < body.getNumChildren(); i++) {
        var child = body.getChild(i);
        if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
        var table = child.asTable();
        for (var r = 0; r < table.getNumRows(); r++) {
            var row = table.getRow(r);
            for (var c = 0; c < row.getNumCells(); c++) {
                var cell = row.getCell(c);
                var text = normalizeCellText(cell.getText());
                for (var j = 0; j < normalizedTargets.length; j++) {
                    if (text === normalizedTargets[j]) {
                        if (c + 1 < row.getNumCells()) {
                            return row.getCell(c + 1);
                        }
                    }
                }
            }
        }
    }
    return null;
}

function normalizeCellText(text) {
    return String(text || "")
        .replace(/\s+/g, "")
        .replace(/[:\uFF1A]/g, "")
        .trim();
}

function resetApprovalSlotByTokens(body, slotIndex) {
    var nameToken = "{{APPR" + slotIndex + "_NAME}}";
    var stampToken = "{{APPR" + slotIndex + "_STAMP}}";
    var applied = false;

    var nameRange = body.findText(escapeRegExp(nameToken));
    if (nameRange) {
        var nameCell = findParentTableCell(nameRange.getElement().asText());
        if (nameCell) {
            setCellText(nameCell, nameToken);
            applied = true;
        }
    }

    var stampRange = body.findText(escapeRegExp(stampToken));
    if (stampRange) {
        var stampCell = findParentTableCell(stampRange.getElement().asText());
        if (stampCell) {
            setCellText(stampCell, stampToken);
            applied = true;
        }
    }

    return applied;
}

function applyApprovalSlotToDocument(body, slot) {
    var slotIndex = Number(slot.slot_index || slot.slot || slot.index || 0);
    if (!slotIndex) throw new Error("slot_index is invalid");

    var name = String(slot.name || "").trim();
    if (!name) return;
    var stampUrl = String(slot.stamp_url || slot.stampUrl || "").trim();
    var slotStatus = String(slot.status || "").trim().toLowerCase();
    var slotKind = String(slot.kind || "").trim().toLowerCase();
    var stampToken = "{{APPR" + slotIndex + "_STAMP}}";
    var expectedTotalSlots = Number(slot.total_slots || slot.totalSlots || 0);
    if (!expectedTotalSlots || expectedTotalSlots < slotIndex) {
        expectedTotalSlots = Math.max(2, slotIndex);
    }
    var layout = locateApprovalTableLayout(body, expectedTotalSlots);
    var layoutNameCell = getNameCellFromLayout(layout, slotIndex);

    var nameCell = null;
    var nameCandidates = ["{{APPR" + slotIndex + "_NAME}}", "\uACB0\uC7AC\uC790" + slotIndex];
    var nameReplace = replaceFirstTextCandidate(body, nameCandidates, name);
    if (nameReplace && nameReplace.replaced) {
        nameCell = nameReplace.cell;
    } else if (layoutNameCell) {
        setCellText(layoutNameCell, name);
        nameCell = layoutNameCell;
    } else {
        // Last resort: locate the exact name in a table cell only (avoid matching body metadata text).
        var fallbackNameCell = findTableCellByExactText(body, name);
        if (fallbackNameCell) {
            setCellText(fallbackNameCell, name);
            nameCell = fallbackNameCell;
        } else {
            throw new Error("name placeholder not found for slot " + slotIndex);
        }
    }

    var stampInserted = false;
    var stampMethod = "";
    var stampCell = getStampCellFromSlot(body, slotIndex, nameCell, layout);

    if (slotKind === "drafter") {
        if (stampUrl) {
            stampInserted = writeStampToCell(stampCell, stampUrl);
            stampMethod = stampCell.method;
        } else {
            clearStampCellByContext(stampCell);
        }
    } else if (slotStatus === "approved") {
        if (stampUrl) {
            stampInserted = writeStampToCell(stampCell, stampUrl);
            stampMethod = stampCell.method;
        } else {
            clearStampCellByContext(stampCell);
        }
    } else if (slotStatus === "rejected") {
        stampInserted = writeRejectTextToCell(stampCell);
        stampMethod = stampCell.method || "text";
    } else {
        restoreStampPlaceholderInCell(stampCell, stampToken);
        stampMethod = (stampCell && stampCell.method) ? (stampCell.method + "_placeholder") : "placeholder";
    }

    return {
        slot_index: slotIndex,
        name_replaced: true,
        status: slotStatus || slotKind || null,
        stamp_requested: !!stampUrl || slotStatus === "rejected",
        stamp_inserted: stampInserted,
        stamp_method: stampMethod || null
    };
}

function findTableCellByExactText(body, targetText) {
    var normalizedTarget = normalizeCellText(targetText);
    if (!normalizedTarget) return null;
    for (var i = 0; i < body.getNumChildren(); i++) {
        var child = body.getChild(i);
        if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
        var table = child.asTable();
        if (isExcludedApprovalTargetTable(table)) continue;
        for (var r = 0; r < table.getNumRows(); r++) {
            var row = table.getRow(r);
            for (var c = 0; c < row.getNumCells(); c++) {
                var cell = row.getCell(c);
                if (normalizeCellText(cell.getText()) === normalizedTarget) {
                    return cell;
                }
            }
        }
    }
    return null;
}

function getStampCellFromSlot(body, slotIndex, nameCell, layout) {
    var token = "{{APPR" + slotIndex + "_STAMP}}";
    var range = body.findText(escapeRegExp(token));
    if (range) {
        var textEl = range.getElement().asText();
        var cell = findParentTableCell(textEl);
        if (cell) {
            return { cell: cell, method: "token", tokenRange: range };
        }
    }
    if (nameCell) {
        var pos = getTableCellPosition(nameCell);
        if (pos && pos.rowIndex + 1 < pos.table.getNumRows()) {
            var row = pos.table.getRow(pos.rowIndex + 1);
            if (pos.colIndex < row.getNumCells()) {
                return { cell: row.getCell(pos.colIndex), method: "below_cell", tokenRange: null };
            }
        }
    }
    if (layout) {
        var slotCell = getStampCellFromLayout(layout, slotIndex);
        if (slotCell) {
            return { cell: slotCell, method: "layout", tokenRange: null };
        }
    }
    throw new Error("stamp cell/token not found");
}

function clearStampCellByContext(ctx) {
    if (!ctx || !ctx.cell) return false;
    if (ctx.tokenRange) {
        var textEl = ctx.tokenRange.getElement().asText();
        textEl.deleteText(ctx.tokenRange.getStartOffset(), ctx.tokenRange.getEndOffsetInclusive());
    }
    clearContainer(ctx.cell);
    return true;
}

function writeStampToCell(ctx, stampUrl) {
    clearStampCellByContext(ctx);
    appendStampImageToCell(ctx.cell, stampUrl);
    return true;
}

function writeRejectTextToCell(ctx) {
    clearStampCellByContext(ctx);
    var p = ctx.cell.appendParagraph("\uBC18\uB824");
    try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) {}
    try {
        var t = p.editAsText();
        t.setForegroundColor("#dc2626");
        t.setBold(true);
    } catch (e) {}
    return true;
}

function restoreStampPlaceholderInCell(ctx, token) {
    if (!ctx || !ctx.cell) return false;
    setCellText(ctx.cell, token || "");
    return true;
}

function getDirectBodyChildIndex(body, element) {
    if (!body || !element) return -1;
    try {
        var cur = element;
        while (cur) {
            var parent = cur.getParent && cur.getParent();
            if (!parent) return -1;
            if (parent.getType && parent.getType() === DocumentApp.ElementType.BODY_SECTION) {
                return body.getChildIndex(cur);
            }
            cur = parent;
        }
    } catch (e) {}
    return -1;
}

function findApprovalAnchoredTable(body) {
    if (!body) return null;
    var startToken = "{{APPROVAL_BLOCK_START}}";
    var endToken = "{{APPROVAL_BLOCK_END}}";
    var startRange = body.findText(escapeRegExp(startToken));
    var endRange = body.findText(escapeRegExp(endToken));
    if (!startRange || !endRange) return null;

    var startEl = startRange.getElement();
    var endEl = endRange.getElement();
    var startIdx = getDirectBodyChildIndex(body, startEl);
    var endIdx = getDirectBodyChildIndex(body, endEl);
    if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx) return null;

    var candidateTable = null;
    var candidateTableIdx = -1;
    for (var i = startIdx + 1; i < endIdx; i++) {
        var child = body.getChild(i);
        if (child.getType() === DocumentApp.ElementType.TABLE) {
            candidateTable = child.asTable();
            candidateTableIdx = i;
            break;
        }
        // Allow only empty-ish separators between anchor and table for stability.
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            var txt = normalizeCellText(child.asParagraph().getText());
            if (!txt) continue;
        }
        // Non-table content between anchors is a strong signal of malformed template.
        if (child.getType() !== DocumentApp.ElementType.TABLE) {
            // keep scanning but mark as malformed by rejecting if no clear table appears
            continue;
        }
    }
    if (!candidateTable) return null;

    return {
        table: candidateTable,
        tableBodyChildIndex: candidateTableIdx,
        startBodyChildIndex: startIdx,
        endBodyChildIndex: endIdx
    };
}

function locateApprovalTableLayoutByAnchors(body, preferredSlots) {
    var anchored = findApprovalAnchoredTable(body);
    if (!anchored || !anchored.table) return null;

    // With anchors, treat the table as authoritative even if it would otherwise look "excluded".
    var best = findBestApprovalLayoutInTable(anchored.table, preferredSlots, false) ||
               findBestApprovalLayoutInTable(anchored.table, preferredSlots, true);
    if (!best) return null;

    return {
        table: best.table,
        nameRowIndex: best.nameRowIndex,
        stampRowIndex: best.stampRowIndex,
        baseColIndex: best.baseColIndex,
        totalSlots: best.totalSlots,
        anchored: true,
        anchor_start_index: anchored.startBodyChildIndex,
        anchor_end_index: anchored.endBodyChildIndex
    };
}

function locateApprovalTableLayout(body, totalSlots) {
    var anchoredLayout = locateApprovalTableLayoutByAnchors(body, totalSlots);
    if (anchoredLayout) return anchoredLayout;

    function buildLayoutFromPos(pos, slotOffset, nameRowIndex, stampRowIndex) {
        if (!pos || !pos.table) return null;
        if (isExcludedApprovalTargetTable(pos.table)) return null;
        if (nameRowIndex < 0 || stampRowIndex < 0) return null;
        if (nameRowIndex >= pos.table.getNumRows() || stampRowIndex >= pos.table.getNumRows()) return null;
        var baseCol = pos.colIndex - slotOffset;
        if (baseCol < 0) return null;
        var nameRow = pos.table.getRow(nameRowIndex);
        var stampRow = pos.table.getRow(stampRowIndex);
        if (baseCol + totalSlots - 1 >= nameRow.getNumCells()) return null;
        if (baseCol + totalSlots - 1 >= stampRow.getNumCells()) return null;
        if (!isLikelyApprovalSlotBlock(pos.table, nameRowIndex, stampRowIndex, baseCol, totalSlots)) return null;
        return { table: pos.table, nameRowIndex: nameRowIndex, stampRowIndex: stampRowIndex, baseColIndex: baseCol };
    }

    for (var i = 1; i <= totalSlots; i++) {
        var nameToken = "{{APPR" + i + "_NAME}}";
        var stampToken = "{{APPR" + i + "_STAMP}}";

        var nameRange = body.findText(escapeRegExp(nameToken));
        if (nameRange) {
            var nameCell = findParentTableCell(nameRange.getElement().asText());
            var pos = nameCell ? getTableCellPosition(nameCell) : null;
            var fromName = buildLayoutFromPos(pos, i - 1, pos ? pos.rowIndex : -1, pos ? (pos.rowIndex + 1) : -1);
            if (fromName) return fromName;
        }

        var stampRange = body.findText(escapeRegExp(stampToken));
        if (stampRange) {
            var stampCell = findParentTableCell(stampRange.getElement().asText());
            var spos = stampCell ? getTableCellPosition(stampCell) : null;
            var fromStamp = buildLayoutFromPos(spos, i - 1, spos ? (spos.rowIndex - 1) : -1, spos ? spos.rowIndex : -1);
            if (fromStamp) return fromStamp;
        }
    }

    var headerLabels = ["\uB2F4\uB2F9", "\uD300\uC7A5", "\uC0AC\uBB34\uAD6D\uC7A5", "\uAD00\uC7A5", "\uACB0\uC7AC", "\uC6D0\uC7A5"];
    for (var b = 0; b < body.getNumChildren(); b++) {
        var child = body.getChild(b);
        if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
        var table = child.asTable();
        if (isExcludedApprovalTargetTable(table)) continue;
        for (var r = 0; r < table.getNumRows(); r++) {
            if (r + 2 >= table.getNumRows()) continue;
            var row = table.getRow(r);
            if (row.getNumCells() < totalSlots) continue;
            var matched = 0;
            for (var c = 0; c < Math.min(row.getNumCells(), headerLabels.length); c++) {
                var t = normalizeCellText(row.getCell(c).getText());
                if (!t) continue;
                for (var h = 0; h < headerLabels.length; h++) {
                    if (t === normalizeCellText(headerLabels[h])) { matched++; break; }
                }
            }
            if (matched >= Math.min(2, totalSlots) && isLikelyApprovalSlotBlock(table, r + 1, r + 2, 0, totalSlots)) {
                return { table: table, nameRowIndex: r + 1, stampRowIndex: r + 2, baseColIndex: 0 };
            }
        }
    }

    var best = null;
    for (var tb = 0; tb < body.getNumChildren(); tb++) {
        var tChild = body.getChild(tb);
        if (tChild.getType() !== DocumentApp.ElementType.TABLE) continue;
        var t = tChild.asTable();
        if (isExcludedApprovalTargetTable(t)) continue;
        // Do not restrict by document position; some templates place approval blocks near the bottom.
        for (var nr = 0; nr < t.getNumRows() - 1; nr++) {
            var nameRow = t.getRow(nr);
            var stampRow = t.getRow(nr + 1);
            var cols = Math.min(nameRow.getNumCells(), stampRow.getNumCells());
            if (cols < totalSlots) continue;
            for (var baseCol = 0; baseCol + totalSlots <= cols; baseCol++) {
                if (!isLikelyApprovalSlotBlock(t, nr, nr + 1, baseCol, totalSlots)) continue;
                var score = 0;
                for (var cc = 0; cc < totalSlots; cc++) {
                    var nText = normalizeCellText(nameRow.getCell(baseCol + cc).getText());
                    var sCell = stampRow.getCell(baseCol + cc);
                    var sText = normalizeCellText(sCell.getText());
                    if (nText && nText.length >= 2) score += 1;
                    if (sText.indexOf("APPR") >= 0) score += 2;
                    if (sText.indexOf("\uBC18\uB824") >= 0) score += 2;
                    if (cellHasInlineImage(sCell)) score += 2;
                }
                if (!best || score > best.score) {
                    best = { score: score, table: t, nameRowIndex: nr, stampRowIndex: nr + 1, baseColIndex: baseCol };
                }
            }
        }
    }
    if (best && best.score >= Math.max(3, totalSlots)) {
        return { table: best.table, nameRowIndex: best.nameRowIndex, stampRowIndex: best.stampRowIndex, baseColIndex: best.baseColIndex };
    }
    return null;
}
function isExcludedApprovalTargetTable(table) {
    if (!table) return false;
    var exact = {};
    exact["\uC2DC\uD589"] = true;
    exact["\uC2DC\uD589\uC77C"] = true;
    exact["\uB300\uD45C\uBA54\uC77C"] = true;
    exact["\uB300\uD45C\uBC88\uD638"] = true;
    exact["\uC6B0"] = true;
    exact["\uD734\uAC00\uAE30\uAC04"] = true;
    exact["\uB300\uC9C1\uC790"] = true;
    exact["\uB300\uC9C1\uC790\uC5C5\uBB34\uB0B4\uC6A9"] = true;
    var contains = ["\uC2DC\uD589", "\uB300\uD45C\uBA54\uC77C", "\uB300\uD45C\uBC88\uD638", "web@"];
    try {
        for (var r = 0; r < table.getNumRows(); r++) {
            var row = table.getRow(r);
            for (var c = 0; c < row.getNumCells(); c++) {
                var txt = normalizeCellText(row.getCell(c).getText());
                if (!txt) continue;
                if (exact[txt]) return true;
                for (var i = 0; i < contains.length; i++) {
                    if (txt.indexOf(contains[i]) >= 0) return true;
                }
            }
        }
    } catch (e) {}
    return false;
}

function scoreApprovalSlotBlock(table, nameRowIndex, stampRowIndex, baseColIndex, totalSlots) {
    if (!table) return -1;
    try {
        if (!isLikelyApprovalSlotBlock(table, nameRowIndex, stampRowIndex, baseColIndex, totalSlots)) return -1;
        var nameRow = table.getRow(nameRowIndex);
        var stampRow = table.getRow(stampRowIndex);
        var score = 0;
        for (var i = 0; i < totalSlots; i++) {
            var nText = normalizeCellText(nameRow.getCell(baseColIndex + i).getText());
            var sCell = stampRow.getCell(baseColIndex + i);
            var sText = normalizeCellText(sCell.getText());
            if (nText && nText.length >= 2) score += 1;
            if (nText.indexOf("APPR") >= 0) score += 2;
            if (sText.indexOf("APPR") >= 0) score += 2;
            if (sText.indexOf("\uBC18\uB824") >= 0) score += 2;
            if (cellHasInlineImage(sCell)) score += 2;
        }
        return score;
    } catch (e) {
        return -1;
    }
}

function findBestApprovalLayoutInTable(table, preferredSlots, allowExcluded) {
    if (!table) return null;
    if (!allowExcluded && isExcludedApprovalTargetTable(table)) return null;
    if (allowExcluded && !isExcludedApprovalTargetTable(table)) return null;

    var candidates = [];
    if (preferredSlots && preferredSlots > 0) candidates.push(preferredSlots);
    for (var c = 1; c <= 6; c++) {
        if (candidates.indexOf(c) < 0) candidates.push(c);
    }

    var best = null;
    for (var r = 0; r < table.getNumRows() - 1; r++) {
        var nameRow = table.getRow(r);
        var stampRow = table.getRow(r + 1);
        var cols = Math.min(nameRow.getNumCells(), stampRow.getNumCells());
        if (cols < 1) continue;
        for (var ci = 0; ci < candidates.length; ci++) {
            var slots = candidates[ci];
            if (slots < 1 || slots > cols) continue;
            for (var baseCol = 0; baseCol + slots <= cols; baseCol++) {
                var score = scoreApprovalSlotBlock(table, r, r + 1, baseCol, slots);
                if (score < Math.max(2, slots)) continue;
                if (!best || score > best.score) {
                    best = {
                        score: score,
                        table: table,
                        nameRowIndex: r,
                        stampRowIndex: r + 1,
                        baseColIndex: baseCol,
                        totalSlots: slots
                    };
                }
            }
        }
    }
    return best;
}

function locateApprovalTableLayoutInExcludedTables(body, totalSlots) {
    var best = null;
    for (var i = 0; i < body.getNumChildren(); i++) {
        var child = body.getChild(i);
        if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
        var table = child.asTable();
        var cand = findBestApprovalLayoutInTable(table, totalSlots, true);
        if (cand && (!best || cand.score > best.score)) best = cand;
    }
    if (!best) return null;
    return {
        table: best.table,
        nameRowIndex: best.nameRowIndex,
        stampRowIndex: best.stampRowIndex,
        baseColIndex: best.baseColIndex,
        totalSlots: best.totalSlots,
        from_excluded_table: true
    };
}

function isLikelyApprovalSlotBlock(table, nameRowIndex, stampRowIndex, baseColIndex, totalSlots) {
    if (!table) return false;
    if (nameRowIndex < 0 || stampRowIndex < 0) return false;
    if (nameRowIndex >= table.getNumRows() || stampRowIndex >= table.getNumRows()) return false;
    var nameRow = table.getRow(nameRowIndex);
    var stampRow = table.getRow(stampRowIndex);
    if (baseColIndex < 0) return false;
    if (baseColIndex + totalSlots - 1 >= nameRow.getNumCells()) return false;
    if (baseColIndex + totalSlots - 1 >= stampRow.getNumCells()) return false;
    var approvalSignals = 0;
    var blockedSignals = 0;
    for (var c = 0; c < totalSlots; c++) {
        var nText = normalizeCellText(nameRow.getCell(baseColIndex + c).getText());
        var sCell = stampRow.getCell(baseColIndex + c);
        var sText = normalizeCellText(sCell.getText());
        if (nText && nText.length <= 12) approvalSignals += 1;
        if (sText.indexOf('APPR') >= 0) approvalSignals += 2;
        if (sText.indexOf('\uBC18\uB824') >= 0) approvalSignals += 2;
        if (cellHasInlineImage(sCell)) approvalSignals += 2;
        if (nText.indexOf('@') >= 0 || sText.indexOf('@') >= 0) blockedSignals += 2;
        if (nText.indexOf('\uC2DC\uD589') >= 0 || sText.indexOf('\uC2DC\uD589') >= 0) blockedSignals += 2;
        if (nText.indexOf('\uB300\uD45C') >= 0 || sText.indexOf('\uB300\uD45C') >= 0) blockedSignals += 2;
        if (nText === '\uC6B0' || sText === '\uC6B0') blockedSignals += 1;
    }
    return approvalSignals >= Math.max(2, totalSlots) && blockedSignals === 0;
}

function rowContainsApprovalArtifacts(row, startCol, count) {
    if (!row) return false;
    var end = Math.min(row.getNumCells(), startCol + count);
    for (var c = startCol; c < end; c++) {
        var cell = row.getCell(c);
        var t = normalizeCellText(cell.getText());
        if (t.indexOf("APPR") >= 0) return true;
        if (t.indexOf("\uBC18\uB824") >= 0) return true;
        if (cellHasInlineImage(cell)) return true;
    }
    return false;
}

function rowLooksLikeApprovalNameRow(row, startCol, count) {
    if (!row) return false;
    var end = Math.min(row.getNumCells(), startCol + count);
    var shortTextCount = 0;
    var blocked = 0;
    for (var c = startCol; c < end; c++) {
        var t = normalizeCellText(row.getCell(c).getText());
        if (!t) continue;
        if (t.indexOf("@") >= 0 || t.indexOf("web@") >= 0) blocked += 2;
        if (t.indexOf("\uC2DC\uD589") >= 0) blocked += 2;
        if (t.indexOf("\uB300\uD45C") >= 0) blocked += 2;
        if (t === "\uC6B0") blocked += 1;
        if (t.indexOf("APPR") >= 0) shortTextCount += 2;
        else if (t.length <= 12) shortTextCount += 1;
    }
    return shortTextCount >= 2 && blocked === 0;
}

function cleanupMisplacedApprovalRows(body, maxSlots) {
    var cleaned = [];
    if (!maxSlots || maxSlots < 2) maxSlots = 4;

    for (var bi = 0; bi < body.getNumChildren(); bi++) {
        var child = body.getChild(bi);
        if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
        var table = child.asTable();
        if (!isExcludedApprovalTargetTable(table)) continue; // only clean obvious non-approval tables
        var preserve = findBestApprovalLayoutInTable(table, Math.min(maxSlots, 6), true);
        var preserveRows = {};
        if (preserve) {
            preserveRows[preserve.nameRowIndex] = true;
            preserveRows[preserve.stampRowIndex] = true;
        }

        // In excluded tables (issue/footer/leave tables), any approval artifacts are invalid.
        // Remove artifact rows and adjacent likely name rows.
        var removeMap = {};
        for (var r = 0; r < table.getNumRows(); r++) {
            if (preserveRows[r]) continue;
            var row = table.getRow(r);
            if (!rowContainsApprovalArtifacts(row, 0, row.getNumCells())) continue;
            removeMap[r] = true;
            if (r > 0) {
                var prev = table.getRow(r - 1);
                if (!preserveRows[r - 1] && rowLooksLikeApprovalNameRow(prev, 0, prev.getNumCells())) removeMap[r - 1] = true;
            }
            if (r + 1 < table.getNumRows()) {
                var next = table.getRow(r + 1);
                if (!preserveRows[r + 1] && rowLooksLikeApprovalNameRow(next, 0, next.getNumCells())) removeMap[r + 1] = true;
            }
        }

        var rowsToRemove = [];
        for (var key in removeMap) {
            if (removeMap.hasOwnProperty(key)) rowsToRemove.push(Number(key));
        }
        rowsToRemove.sort(function(a, b) { return b - a; });
        for (var i = 0; i < rowsToRemove.length; i++) {
            var rowIndex = rowsToRemove[i];
            if (rowIndex < 0 || rowIndex >= table.getNumRows()) continue;
            try {
                table.removeRow(rowIndex);
                cleaned.push({ table_child_index: bi, row_index: rowIndex });
            } catch (e) {
                // best effort
            }
        }
    }
    return cleaned;
}

function getDefaultApprovalHeaderLabels(totalSlots) {
    var common = [
        "\uB2F4\uB2F9",
        "\uD300\uC7A5",
        "\uC0AC\uBB34\uAD6D\uC7A5",
        "\uAD00\uC7A5"
    ];
    var labels = [];
    for (var i = 0; i < totalSlots; i++) {
        labels.push(common[i] || ("\uACB0\uC7AC" + (i + 1)));
    }
    return labels;
}

function createFallbackApprovalTableLayout(body, totalSlots) {
    totalSlots = Number(totalSlots || 4);
    if (!totalSlots || totalSlots < 1) totalSlots = 4;
    if (totalSlots > 8) totalSlots = 8;
    var headers = getDefaultApprovalHeaderLabels(totalSlots);
    var nameRow = [];
    var stampRow = [];
    for (var i = 1; i <= totalSlots; i++) {
        nameRow.push("{{APPR" + i + "_NAME}}");
        stampRow.push("{{APPR" + i + "_STAMP}}");
    }
    var rows = [headers, nameRow, stampRow];
    var insertIndex = 0;
    try {
        for (var b = 0; b < body.getNumChildren(); b++) {
            var child = body.getChild(b);
            var t = child.getType();
            if (t === DocumentApp.ElementType.TABLE && !isExcludedApprovalTargetTable(child.asTable())) {
                insertIndex = b;
                break;
            }
            if (t === DocumentApp.ElementType.PARAGRAPH) {
                var txt = normalizeCellText(child.asParagraph().getText());
                if (txt && txt.length > 0) {
                    insertIndex = b;
                    break;
                }
            }
        }
    } catch (e) {}

    var table = body.insertTable(insertIndex, rows);
    try {
        for (var r = 0; r < table.getNumRows(); r++) {
            var row = table.getRow(r);
            for (var c = 0; c < row.getNumCells(); c++) {
                var cell = row.getCell(c);
                for (var p = 0; p < cell.getNumChildren(); p++) {
                    var ch = cell.getChild(p);
                    if (ch.getType && ch.getType() === DocumentApp.ElementType.PARAGRAPH) {
                        ch.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                    }
                }
            }
        }
    } catch (e) {}

    return { table: table, nameRowIndex: 1, stampRowIndex: 2, baseColIndex: 0, inserted_fallback: true };
}
function cellHasInlineImage(cell) {
    if (!cell) return false;
    try {
        for (var i = 0; i < cell.getNumChildren(); i++) {
            var child = cell.getChild(i);
            if (child.getType && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                var p = child.asParagraph();
                for (var j = 0; j < p.getNumChildren(); j++) {
                    var pc = p.getChild(j);
                    if (pc.getType && pc.getType() === DocumentApp.ElementType.INLINE_IMAGE) return true;
                }
            }
        }
    } catch (e) {}
    return false;
}

function getNameCellFromLayout(layout, slotIndex) {
    if (!layout || !layout.table) return null;
    var colIndex = layout.baseColIndex + (slotIndex - 1);
    if (colIndex < 0) return null;
    if (layout.nameRowIndex < 0 || layout.nameRowIndex >= layout.table.getNumRows()) return null;
    var row = layout.table.getRow(layout.nameRowIndex);
    if (colIndex >= row.getNumCells()) return null;
    return row.getCell(colIndex);
}

function getStampCellFromLayout(layout, slotIndex) {
    if (!layout || !layout.table) return null;
    var colIndex = layout.baseColIndex + (slotIndex - 1);
    if (colIndex < 0) return null;
    if (layout.stampRowIndex < 0 || layout.stampRowIndex >= layout.table.getNumRows()) return null;
    var row = layout.table.getRow(layout.stampRowIndex);
    if (colIndex >= row.getNumCells()) return null;
    return row.getCell(colIndex);
}

function setCellText(cell, text) {
    clearContainer(cell);
    var p = cell.appendParagraph(String(text || ""));
    try { p.setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch (e) {}
    return true;
}

function deleteDriveFile(data) {
    var docId = String(data.doc_id || data.docId || "").trim();
    if (!docId) throw new Error("doc_id is required");
    var file = DriveApp.getFileById(docId);
    file.setTrashed(true);
    var trashed = true;
    try {
        trashed = file.isTrashed();
    } catch (e) {
        // Some contexts may not expose isTrashed reliably; assume setTrashed succeeded if no exception.
        trashed = true;
    }
    if (!trashed) {
        throw new Error("Drive file was not moved to trash");
    }
    return {
        doc_id: docId,
        trashed: true,
        file_name: file.getName()
    };
}

function moveDriveFile(data) {
    var docId = String(data.doc_id || data.docId || "").trim();
    if (!docId) throw new Error("doc_id is required");
    var targetFolderId = String(data.target_folder_id || data.targetFolderId || "").trim() || DELETED_ARCHIVE_FOLDER_ID;
    if (!targetFolderId) throw new Error("target_folder_id is required");

    var file = DriveApp.getFileById(docId);
    var folder = DriveApp.getFolderById(targetFolderId);
    file.moveTo(folder);

    return {
        doc_id: docId,
        target_folder_id: targetFolderId,
        file_name: file.getName(),
        web_view_url: "https://drive.google.com/file/d/" + docId + "/view"
    };
}

function createPdfSnapshot(data) {
    var docId = String(data.doc_id || data.docId || "").trim();
    if (!docId) throw new Error("doc_id is required");
    var snapshotName = String(data.snapshot_name || data.snapshotName || "").trim();
    var replaceFileId = String(data.replace_file_id || data.replaceFileId || "").trim();

    var source = DriveApp.getFileById(docId);
    var pdfBlob = null;
    var usedFlattenedDoc = false;
    var exportMethod = "direct";

    try {
        var sourceDoc = DocumentApp.openById(docId);
        var exportUrl = "https://docs.google.com/document/d/" + docId + "/export?format=pdf";
        try {
            if (sourceDoc.getTabs) {
                var tabs = sourceDoc.getTabs();
                if (tabs && tabs.length > 0 && tabs[0].getId) {
                    exportUrl += "&tab=" + encodeURIComponent(tabs[0].getId());
                    exportMethod = "tab_export";
                }
            }
        } catch (tabErr) {
            Logger.log("createPdfSnapshot tab export fallback: " + tabErr.message);
        }
        var resp = UrlFetchApp.fetch(exportUrl, {
            headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
        });
        if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
            throw new Error("PDF export HTTP " + resp.getResponseCode());
        }
        pdfBlob = resp.getBlob();
    } catch (e) {
        Logger.log("createPdfSnapshot export fallback: " + e.message);
        // Final fallback for compatibility.
        pdfBlob = source.getAs(MimeType.PDF);
        exportMethod = "drive_getAs";
    }

    if (!pdfBlob) throw new Error("pdf export failed");
    if (snapshotName) {
        try { pdfBlob.setName(snapshotName); } catch (e) {}
    } else {
        try { pdfBlob.setName(source.getName() + "_print.pdf"); } catch (e) {}
    }

    var targetFolder = null;
    try {
        targetFolder = DriveApp.getFolderById(PRINT_SNAPSHOT_FOLDER_ID);
    } catch (e) {
        Logger.log("createPdfSnapshot fixed print folder fallback: " + e.message);
    }
    if (!targetFolder) {
        try {
            var parents = source.getParents();
            if (parents.hasNext()) {
                targetFolder = parents.next();
            }
        } catch (e) {}
    }
    if (!targetFolder) {
        throw new Error("source parent folder not found");
    }

    if (replaceFileId) {
        try {
            var oldFile = DriveApp.getFileById(replaceFileId);
            oldFile.setTrashed(true);
        } catch (e) {
            Logger.log("createPdfSnapshot replace cleanup skipped: " + e.message);
        }
    }

    var pdfFile = targetFolder.createFile(pdfBlob);
    if (snapshotName) {
        try { pdfFile.setName(snapshotName); } catch (e) {}
    }

    return {
        file_id: pdfFile.getId(),
        id: pdfFile.getId(),
        name: pdfFile.getName(),
        mime_type: pdfFile.getMimeType(),
        size: pdfFile.getSize(),
        web_view_url: pdfFile.getUrl(),
        source_doc_id: docId,
        used_flattened_doc: usedFlattenedDoc,
        export_method: exportMethod
    };
}

function readDriveFileBase64(data) {
    var fileId = String(data.file_id || data.fileId || data.doc_id || data.docId || "").trim();
    if (!fileId) throw new Error("file_id is required");
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return {
        file_id: file.getId(),
        name: file.getName(),
        mime_type: blob.getContentType() || file.getMimeType() || "application/octet-stream",
        size: blob.getBytes().length,
        data_base64: Utilities.base64Encode(blob.getBytes())
    };
}

function copyDocumentBodyForPrint(sourceBody, targetBody) {
    while (targetBody.getNumChildren() > 0) {
        targetBody.removeChild(targetBody.getChild(0));
    }
    for (var i = 0; i < sourceBody.getNumChildren(); i++) {
        appendBodyChildCopy(targetBody, sourceBody.getChild(i));
    }
}

function stripLeadingPrintArtifacts(body) {
    var removed = 0;
    while (body.getNumChildren() > 0) {
        var child = body.getChild(0);
        var type = child.getType();
        if (type === DocumentApp.ElementType.PAGE_BREAK) {
            body.removeChild(child);
            removed++;
            continue;
        }
        if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
            var text = "";
            try { text = String(child.asParagraph ? child.asParagraph().getText() : child.getText() || ""); } catch (e) { text = ""; }
            // Remove leading blank/garbage lines ("??" residues, quote marks, whitespace)
            var normalized = text
                .replace(/\s+/g, "")
                .replace(/[??꺜?봔?떷?'`.,夷??◈?-_=+~]/g, "")
                .trim();
            if (!normalized || normalized === "??") {
                body.removeChild(child);
                removed++;
                continue;
            }
        }
        break;
    }
    return removed;
}

function appendBodyChildCopy(targetBody, child) {
    var type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
        targetBody.appendParagraph(child.copy().asParagraph());
        return;
    }
    if (type === DocumentApp.ElementType.TABLE) {
        targetBody.appendTable(child.copy().asTable());
        return;
    }
    if (type === DocumentApp.ElementType.LIST_ITEM) {
        targetBody.appendListItem(child.copy().asListItem());
        return;
    }
    if (type === DocumentApp.ElementType.PAGE_BREAK) {
        try {
            targetBody.appendPageBreak(child.copy().asPageBreak());
        } catch (e) {
            targetBody.appendPageBreak();
        }
        return;
    }
    if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
        try {
            targetBody.appendHorizontalRule();
        } catch (e) {
            targetBody.appendParagraph("");
        }
        return;
    }
    // Ignore unsupported container types for print snapshot; most approval templates use paragraphs/tables.
}

function placeStampAtToken(body, tokenCandidates, stampUrl) {
    for (var i = 0; i < tokenCandidates.length; i++) {
        var candidate = tokenCandidates[i];
        var range = body.findText(escapeRegExp(candidate));
        if (!range) continue;
        var textEl = range.getElement().asText();
        textEl.deleteText(range.getStartOffset(), range.getEndOffsetInclusive());

        var cell = findParentTableCell(textEl);
        if (cell) {
            clearContainer(cell);
            appendStampImageToCell(cell, stampUrl);
            return true;
        }

        var paragraph = findParentParagraph(textEl);
        if (paragraph) {
            var blob = fetchStampBlob(stampUrl);
            if (!blob) return false;
            paragraph.appendInlineImage(blob);
            return true;
        }
    }
    return false;
}

function replaceFirstTextCandidate(body, candidates, replacement) {
    for (var i = 0; i < candidates.length; i++) {
        var candidate = candidates[i];
        var range = body.findText(escapeRegExp(candidate));
        if (!range) continue;

        var textEl = range.getElement().asText();
        var start = range.getStartOffset();
        var end = range.getEndOffsetInclusive();
        textEl.deleteText(start, end);
        textEl.insertText(start, replacement);

        return {
            replaced: true,
            cell: findParentTableCell(textEl)
        };
    }
    return { replaced: false, cell: null };
}

function placeStampInCellBelow(nameCell, stampUrl) {
    var pos = getTableCellPosition(nameCell);
    if (!pos) return false;
    if (pos.rowIndex + 1 >= pos.table.getNumRows()) return false;

    var nextRow = pos.table.getRow(pos.rowIndex + 1);
    if (pos.colIndex >= nextRow.getNumCells()) return false;
    var targetCell = nextRow.getCell(pos.colIndex);

    clearContainer(targetCell);
    appendStampImageToCell(targetCell, stampUrl);
    return true;
}

function appendStampImageToCell(cell, stampUrl) {
    var blob = fetchStampBlob(stampUrl);
    if (!blob) throw new Error("stamp image fetch failed");
    var p = cell.appendParagraph("");
    try {
        p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } catch (e) {
        // ignore alignment failures in older document contexts
    }
    var img = p.appendInlineImage(blob);
    resizeInlineImage(img, APPROVAL_STAMP_MAX_WIDTH, APPROVAL_STAMP_MAX_HEIGHT);
}

function fetchStampBlob(stampUrl) {
    try {
        var resp = UrlFetchApp.fetch(stampUrl, { muteHttpExceptions: true, followRedirects: true });
        var code = resp.getResponseCode();
        if (code < 200 || code >= 300) {
            throw new Error("HTTP " + code);
        }
        return resp.getBlob();
    } catch (e) {
        Logger.log("fetchStampBlob error: " + e.message);
        return null;
    }
}

function resizeInlineImage(img, maxW, maxH) {
    try {
        var w = img.getWidth();
        var h = img.getHeight();
        if (!w || !h) return;
        var ratio = Math.min(maxW / w, maxH / h, 1);
        img.setWidth(Math.round(w * ratio));
        img.setHeight(Math.round(h * ratio));
    } catch (e) {
        Logger.log("resizeInlineImage error: " + e.message);
    }
}

function clearContainer(container) {
    while (container.getNumChildren && container.getNumChildren() > 0) {
        container.removeChild(container.getChild(0));
    }
}

function findParentTableCell(element) {
    var cur = element;
    while (cur) {
        if (cur.getType && cur.getType() === DocumentApp.ElementType.TABLE_CELL) {
            return cur.asTableCell ? cur.asTableCell() : cur;
        }
        cur = cur.getParent ? cur.getParent() : null;
    }
    return null;
}

function findParentParagraph(element) {
    var cur = element;
    while (cur) {
        if (cur.getType && cur.getType() === DocumentApp.ElementType.PARAGRAPH) {
            return cur.asParagraph ? cur.asParagraph() : cur;
        }
        cur = cur.getParent ? cur.getParent() : null;
    }
    return null;
}

function getTableCellPosition(cell) {
    var row = cell.getParent();
    if (!row || row.getType() !== DocumentApp.ElementType.TABLE_ROW) return null;
    var table = row.getParent();
    if (!table || table.getType() !== DocumentApp.ElementType.TABLE) return null;

    var rowIndex = -1;
    for (var r = 0; r < table.getNumRows(); r++) {
        if (table.getRow(r) === row) {
            rowIndex = r;
            break;
        }
    }
    if (rowIndex < 0) return null;

    var colIndex = -1;
    for (var c = 0; c < row.getNumCells(); c++) {
        if (row.getCell(c) === cell) {
            colIndex = c;
            break;
        }
    }
    if (colIndex < 0) return null;

    return { table: table, rowIndex: rowIndex, colIndex: colIndex };
}

function escapeRegExp(text) {
    return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function uploadImageToDrive(imageData, username, kind) {
    try {
        var typeLabel = (kind === "stamp") ? "stamp" : "profile";
        var decoded = Utilities.base64Decode(imageData.data);
        var blob = Utilities.newBlob(decoded, imageData.type || "image/png", imageData.name || (username + "_" + typeLabel + ".png"));

        // Use specific Drive folder for user images
        var folder = DriveApp.getFolderById("1ZwpNkDKZKw2EZ9tqR4RCL4ILiaE6BnaG");

        // Delete old same-type image for this user if exists
        var baseName = username + "_" + typeLabel;
        var existingFiles = folder.getFilesByName(baseName);
        while (existingFiles.hasNext()) {
            existingFiles.next().setTrashed(true);
        }

        // Create new file
        var file = folder.createFile(blob);
        file.setName(baseName);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        // Return direct thumbnail URL for embedding
        var fileId = file.getId();
        return { url: "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w400", error: null };
    } catch (e) {
        Logger.log("Image upload error: " + e.message);
        return { url: null, error: e.message };
    }
}

function getUsers() {
    const ss = SpreadsheetApp.openById(USER_SHEET_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const users = [];

    // Skip header
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        users.push({
            username: row[0],
            full_name: row[2],
            department: row[3],
            job_title: row[4],
            role: row[5],
            total_leave: row[6],
            used_leave: row[7]
        });
    }

    return ContentService.createTextOutput(JSON.stringify({ users: users }))
        .setMimeType(ContentService.MimeType.JSON);
}

function success(data) {
    // Flatten image URLs to top level for server to read
    var response = { status: "success", data: data, script_build: SCRIPT_BUILD };
    if (data && data.profile_image_url) {
        response.profile_image_url = data.profile_image_url;
    }
    if (data && data.approval_stamp_image_url) {
        response.approval_stamp_image_url = data.approval_stamp_image_url;
    }
    return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: msg }))
        .setMimeType(ContentService.MimeType.JSON);
}

function deleteUser(username) {
    const ss = SpreadsheetApp.openById(USER_SHEET_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] == username) {
            sheet.deleteRow(i + 1); // 1-based index
            break;
        }
    }
}

