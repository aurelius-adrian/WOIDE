const todo = "TODO";

const doctitleElement = ({ doctitle }, generator) => `<span shtml:doctitle="${doctitle}" style="display:none" >${generator}</span>`;

const moduleElement = ({
                         document_uri,
                         module_name,
                         language
                       }, generator) => `<span shtml:theory="${document_uri}?${module_name}" shtml:language="${language}" shtml:signature="">${generator}</span>`;

const symbolElement = ({
                         document_uri,
                         module_name,
                         tag_type,
                         symbol_name
                       }, generator) => `<span shtml:${tag_type}="${document_uri}?${module_name}?${symbol_name}">${generator}</span>`;

const referenceElement = ({
                              document_uri,
                              module_name,
                              symbol_name
                          }, generator) =>
`<span shtml:term="OMID" shtml:notationid="" shtml:head="${document_uri}?${module_name}?${symbol_name}">
    <span shtml:comp="${document_uri}?${module_name}?${symbol_name}">${generator}</span>
</span>`;

async function goToReferencedTag(event) {
  Word.run(async (context) => {
    let e = await getTagsByAnnotationType(context, "Reference");
    e = e.find(e => e.id === event.ids[0]);
    let title = e.title.split(":");
    let s = await getTagsByAnnotationType(context, "Symbol");
    s = s.find((e) => {
        let cTitle = e.title.split(":");
        return cTitle[0] === title[0] && cTitle[2] === title[1];
      });
    s.select(Word.SelectionMode.select);
  });
}

const AnnotationTypes = {
  DocumentTitle: {
    name: "DocumentTitle",
    color: "#0e90d2",
    fields: [
      {
        id: "doctitle",
        name: "Document Title",
        default: "",
        options: []
      }
    ],
    export: doctitleElement,
    onEntered: () => {}
  },
  Module: {
    name: "Module",
    color: "#8dd6ff",
    fields: [
      {
        id: "module_name",
        name: "Module Name",
        default: "",
        options: []
      },
      {
        id: "language",
        name: "Language",
        default: "en",
        options: ["de", "en"]
      }
    ],
    export: moduleElement,
    onEntered: () => {}
  },
  Symbol: {
    name: "Symbol",
    color: "#ed028c",
    fields: [
      {
        id: "module_name",
        name: "Module Name",
        default: "",
        options: []
      },
      {
        id: "tag_type",
        name: "Tag Type",
        default: "",
        options: ["Definiens", "Definiendum"]
      },
      {
        id: "symbol_name",
        name: "Symbol Name",
        default: "",
        options: []
      }
    ],
    export: symbolElement,
    onEntered: () => {}
  },
  Reference: {
    name: "Reference",
    color: "#0e90d2",
    fields: [
      {
        id: "module_name",
        name: "Module Name",
        default: "",
        options: []
      },
      {
        id: "symbol_name",
        name: "Symbol Name",
        default: "",
        options: []
      }
    ],
    export: referenceElement,
    onEntered: goToReferencedTag
  }
};

function getAnnotationObject(tag, str) {
  let components = str.split(":");
  if (components.length !== AnnotationTypes[tag].fields.length)
    throw new Error("Structure does not match given annotation type");
  let ret = {};
  for (let i = 0; i < components.length; i++) {
    ret[AnnotationTypes[tag].fields[i].id] = components[i];
  }
  ret.document_uri = Office.context.document.settings.get("Document-URI");
  return ret;
}

function createObject(type, data, generator) {
  return AnnotationTypes[type].export(data, generator);
}

