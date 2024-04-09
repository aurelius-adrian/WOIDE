/* global document, Office, Word */
// ----- Templates -----
const getFields = ({ name, fields }) => {
  let ret = "";
  for (let i = 0; i < fields.length; i++) {
    ret += `<label for="${fields[i].id}">${fields[i].name}</label><br>
      <input type="text" id="${fields[i].id}" name="${fields[i].name}" list="${fields[i].id}_list">
      <datalist id="${fields[i].id}_list">
      ${getDatalist(fields[i])}
      </datalist>
      <br>`;
  }
  return ret;
};

const getDatalist = ({ options }) => {
  let ret = "";
  for (let i = 0; i < options.length; i++) {
    ret += `<option value=\"${options[i]}\" />`;
  }
  return ret;
};

const getAnnotationsTypeOptions = () => {
  let ret = "";
  for (let i = 0; i < Object.keys(AnnotationTypes).length; i++) {
    let tmp = AnnotationTypes[Object.keys(AnnotationTypes)[i]];
    ret += `<option value="${Object.keys(AnnotationTypes)[i]}">${tmp.name}</option>`;
  }
  return ret;
};
// ----- ----- -----

let toggleTypes = [Word.ContentControlAppearance.tags, Word.ContentControlAppearance.boundingBox, Word.ContentControlAppearance.hidden];
let toggleIndex = 0;

let messageBanner;

async function test() {
  Word.run(async (context) => {
    console.log(toggleTypes);
    console.log(toggleIndex);
  });
}

async function getExportHtml(ccs, context, start=true) {
  if (ccs.length === 0) return "";
  let ret = '';
  for (let i = 0; i < ccs.length; i++) {
    let tmp = ccs[i].getHtml();
    await context.sync();
    let content = getHTMLBody(tmp.value);
    if (ccs[i].children.length !== 0) {
      content = await getExportHtml(ccs[i].children, context, false);
    }
    ret += getHTMLBody(ccs[i].htmlBefore.value) + createObject(ccs[i].tag, getAnnotationObject(ccs[i].tag, ccs[i].title),
      content);
  }
  if (start) {
    let head = context.document.body.getHtml();
    await context.sync();
    return "<html>" + getHTMLHead(head.value) + ret + getHTMLBody(ccs[ccs.length - 1].htmlAfter.value) + "</html>";
  }
  return ret + getHTMLBody(ccs[ccs.length - 1].htmlAfter.value);
}

function getHTMLBody(html) {
  const parser = new DOMParser();
  const htmlDoc = parser.parseFromString(html, "text/html");
  return $(".WordSection1", htmlDoc)[0].innerHTML.replace(/\n\n/g, "\n");
}

function getHTMLHead(html) {
  const parser = new DOMParser();
  const htmlDoc = parser.parseFromString(html, "text/html");
  return $("head", htmlDoc)[0].innerHTML.replace(/\n\n/g, "\n");
}

// $("head", new DOMParser().parseFromString(context.document.body.getHtml(), "text/html");)[0].innerHTML

async function getCCs(context, parentCC) {
  let ccs = null;
  // if (parentCC === null) {
  ccs = context.document.contentControls;
  // } else {
  //   ccs = parentCC.contentControls;
  // }

  ccs.load();
  await context.sync();

  if (ccs.items.length === 0) {
    return [];
  }

  let ret = [];

  if (parentCC === null) {
    parentCC = {
      isNullObject: true,
      getRange: function(p) {
        return context.document.body.getRange(p);
      }
    };
  }

  for (let i = 0; i < ccs.items.length; i++) {
    if (!Object.keys(AnnotationTypes).includes(ccs.items[i].tag)) continue;
    let tp = ccs.items[i].parentContentControlOrNullObject;
    tp.load();
    await context.sync();
    if (!tp.isNullObject && !parentCC.isNullObject) {
      if (tp.id !== parentCC.id) continue;
    } else if (tp.isNullObject ^ parentCC.isNullObject) continue;

    let rangeAfter = parentCC.getRange(Word.RangeLocation.end).expandTo(ccs.items[i].getRange(Word.RangeLocation.end));
    let rangeBefore = null;
    rangeAfter.load();
    if (ret.length === 0) {
      rangeBefore = parentCC.getRange(Word.RangeLocation.start).expandTo(ccs.items[i].getRange(Word.RangeLocation.start));
      rangeBefore.load();
      await context.sync();

      ret = [ccs.items[i]];
      ret[0].children = await getCCs(context, ret[0]);

      ret[0].rangeBefore = rangeBefore;
      ret[0].htmlBefore = rangeBefore.getHtml();
      ret[0].rangeAfter = rangeAfter;
      ret[0].htmlAfter = rangeAfter.getHtml();
      await context.sync();

      continue;
    }
    rangeBefore = ret.at(-1).getRange(Word.RangeLocation.end).expandTo(ccs.items[i].getRange(Word.RangeLocation.start));
    rangeBefore.load();
    await context.sync();

    let b = false;
    for (let j = 0; j < ret.length; j++) {
      let pos = ret[j].getRange().compareLocationWith(ccs.items[i].getRange());
      await context.sync();

      if (pos.value === Word.LocationRelation.adjacentAfter || pos.value === Word.LocationRelation.after) {
        ccs.items[i].children = await getCCs(context, ccs.items[i]);

        rangeAfter = ret[j].getRange(Word.RangeLocation.start).expandTo(ccs.items[i].getRange(Word.RangeLocation.end));
        rangeAfter.load();
        await context.sync();

        ccs.items[i].rangeBefore = rangeBefore;
        ccs.items[i].htmlBefore = rangeBefore.getHtml();
        ccs.items[i].rangeAfter = rangeAfter;
        ccs.items[i].htmlAfter = rangeAfter.getHtml();

        ret[j].rangeBefore = rangeAfter;
        ret[j].htmlBefore = rangeAfter.getHtml();

        if (j > 0) {
          ret[j - 1].rangeAfter = rangeBefore;
          ret[j - 1].htmlAfter = rangeBefore.getHtml();
        }
        await context.sync();

        ret.splice(j, 0, ccs.items[i]);
        b = true;
        break;
      }
        // else if (pos.value === Word.LocationRelation.adjacentBefore || pos.value === Word.LocationRelation.before) {
        //   console.log("This shouldn't happen... 0");
      // }
      else if (pos.value === Word.LocationRelation.contains || pos.value === Word.LocationRelation.containsEnd || pos.value === Word.LocationRelation.containsStart) {
        console.log("This shouldn't happen... 1");
        b = true;
        break;
      } else if (pos.value === Word.LocationRelation.inside || pos.value === Word.LocationRelation.insideEnd || pos.value === Word.LocationRelation.insideStart) {
        console.log("This shouldn't happen... 2");
        b = true;
        break;
      }
    }
    if (!b) {
      ccs.items[i].children = await getCCs(context, ccs.items[i]);

      ccs.items[i].rangeBefore = rangeBefore;
      ccs.items[i].htmlBefore = rangeBefore.getHtml();
      ccs.items[i].rangeAfter = rangeAfter;
      ccs.items[i].htmlAfter = rangeAfter.getHtml();

      ret.at(-1).rangeAfter = rangeBefore;
      ret.at(-1).htmlAfter = rangeBefore.getHtml();
      await context.sync();

      ret.push(ccs.items[i]);
    }
  }

  return ret;
}


/*Office.onReady(() => {
    Office.addinCommands
        .addCommand({
            displayName: "My Command",
            icon: "my-icon",
            onExecute: () => {
                // Code to execute when the command is triggered
            },
            keyBindings: [
                {
                    key: "F12"
                },
                {
                    key: "ctrl+shift+m"
                }
            ]
        });
    
    Office.addinCommands
        .getCommand("My Command")
        .then(command => {
            OfficeRuntime.keyboard
                .addShortcut(command.keyBindings[0].key, () => {
                    command.onExecute();
                });
        });
});*/

(function() {
  "use strict";
  Office.initialize = function(reason) {
    $(document).ready(function() {
      let element = document.querySelector(".MessageBanner");
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();

      $("#template-description").text("This sample can be used to tag the selected text.");

      $("#tag-button-text").text("Set Tag");
      $("#tag-button-desc").text("Sets a annotation tag with the entered tag and title.");

      $("#untag-button-text").text("Remove Tags");
      $("#untag-button-desc").text("Removes all annotation tags in the selection.");

      $("#toggle-button-text").text("Toggle Tags");
      $("#toggle-button-desc").text("Toggles visibility of all tags.");

      $("#save-button-text").text("Save");
      $("#save-button-desc").text("Saves all annotations to a HTML document.");

      $("#test-button-text").text("Test");
      $("#test-button-desc").text("Gets all tags in the document and prints their JSON.");

      $("#tag-button").click(insertCC);
      $("#untag-button").click(removeCC);
      $("#toggle-button").click(toggleCC);
      $("#test-button").click(test);
      $("#save-button").click(saveAnnotations);

      console.log(AnnotationTypes);
      $("#form_select").html(getAnnotationsTypeOptions());
      $("#form_select").change(() => {
        $("#form").html(getFields(AnnotationTypes[$("#form_select").val()]));
      });

      let saveTimeout;
      $("#document_uri").on("input", function() {
        saveTimeout = setTimeout(() => setDocumentURI($(this).val()), 100);
      });

      $("#document_uri").val(Office.context.document.settings.get("Document-URI"));

      setOnEnter();
    });
  };
})();

function setDocumentURI(uri) {
  Office.context.document.settings.set("Document-URI", uri);
  Office.context.document.settings.saveAsync();
  showNotification("Settings", "Document URI was saved.");
}

function saveAnnotations() {
  Word.run(async (context) => {
    let ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();
    let ccc = await getCCs(context, null);
    console.log(ccc);
    let out = await getExportHtml(ccc, context);
    console.log(out);

    let currentDate = new Date().toISOString();
    download(out, `export-${currentDate}.html`, "text/plain");
  });
}

function download(data, filename, type) {
  let file = new Blob([data], { type: type });
  if (window.navigator.msSaveOrOpenBlob) // IE10+
    window.navigator.msSaveOrOpenBlob(file, filename);
  else { // Others
    let a = document.createElement("a"),
      url = URL.createObjectURL(file);
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function() {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 0);
  }
}

function setOnEnter() {
  Word.run(async (context) => {
    let ccs = await getTagsByAnnotationType(context, "Reference");
    console.log(ccs);
    for (let i = 0; i < ccs.length; i++) {
      ccs[i].onEntered.add(AnnotationTypes.Reference.onEntered);
      ccs[i].track();
      console.log("test");
    }
    await context.sync();
  });
}

function checkSeparator(str) {
  // return str.includes('\u200b');
  return str.includes(":");
}

function insertCC() {
  Word.run(async (context) => {
    let selection = context.document.getSelection();

    let tag = $("#form_select").val();
    let annotationType = AnnotationTypes[tag];
    let str = "";
    for (let i = 0; i < annotationType.fields.length; i++) {
      let tmp = $(`#${annotationType.fields[i].id}`).val();
      if (checkSeparator(tmp)) {
        showNotification("Tag id may not include ':'");
        return;
      }
      str += tmp + ":";
    }
    str = str.slice(0, -1);

    let ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    let cc = selection.insertContentControl();
    cc.tag = tag;
    cc.title = str;
    cc.appearance = toggleTypes[toggleIndex];
    cc.color = annotationType.color;
    if (tag === "Reference")
      cc.onEntered.add(AnnotationTypes.Reference.onEntered);
      cc.track();
    await context.sync();
  });
}

function removeCC() {
  Word.run(async (context) => {
    let selection = context.document.getSelection();
    let ccs = selection.contentControls;
    ccs.load();
    await context.sync();
    for (let i = 0; i < ccs.items.length; i++) {
      ccs.items[i].delete(true);
    }
  });
}

function toggleCC() {
  Word.run(async (context) => {
    let ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    test();
    console.log(toggleIndex);
    toggleIndex = (toggleIndex + 1) % 3;
    for (let i = 0; i < ccs.items.length; i++) {
      ccs.items[i].appearance = toggleTypes[toggleIndex];
    }
  });
}

async function getTagsByAnnotationType(context, type) {
  let ccs = context.document.contentControls;
  ccs.load("items");
  await context.sync();
  return ccs.items.filter(t => t.tag === type);
}

function getCC() {
  Word.run(async (context) => {
    let ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    let out = "";

    for (let i = 0; i < ccs.items.length; i++) {
      out +=
        "{'tag': '" +
        ccs.items[i].tag +
        "', 'title': '" +
        ccs.items[i].title +
        "', 'id': '" +
        ccs.items[i].id +
        "'}\n";
    }

    showNotification(out);
  });
}

function getSortedCCs() {
  Word.run(async function(context) {
    let ccs = context.document.contentControls;
    ccs.load();
    await context.sync();
  });
}

//$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
function errorHandler(error) {
  // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
  showNotification("Error:", error);
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
  }
}

// Helper function for displaying notifications
function showNotification(header, content) {
  $("#notification-header").text(header);
  $("#notification-body").text(content);
  messageBanner.showBanner();
  messageBanner.toggleExpansion();
}