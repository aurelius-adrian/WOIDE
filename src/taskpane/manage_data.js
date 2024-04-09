/*$("#add-custom-xml-part").click(() => tryCatch(addCustomXmlPart));
$("#query").click(() => tryCatch(query));
$("#insert-attribute").click(() => tryCatch(insertAttribute));
$("#insert-element").click(() => tryCatch(insertElement));
$("#delete-custom-xml-part").click(() => tryCatch(deleteCustomXmlPart));*/

async function addCustomXmlPart() {
    // Add a custom XML part.
    await Word.run(async (context) => {
        const originalXml =
            "<Tags><Tag>Definiens</Tag><Tag>Hong</Tag><Tag>Sally</Tag></Tags>";
        const customXmlPart = context.document.customXmlParts.add(originalXml);
        customXmlPart.load("id");
        const xmlBlob = customXmlPart.getXml();

        await context.sync();

        const readableXml = addLineBreaksToXML(xmlBlob.value);
        console.log("Added custom XML part:");
        console.log(readableXml);

        // Store the XML part's ID in a setting so the ID is available to other functions.
        const settings = context.document.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);

        await context.sync();
    });
}

async function query() {
    // Query a custom XML part for elements matching the search terms.
    await Word.run(async (context) => {
        const settings = context.document.settings;
        const xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

        await context.sync();

        if (xmlPartIDSetting.value) {
            const customXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
            const xpathToQueryFor = "/Reviewers/Reviewer";
            const clientResult = customXmlPart.query(xpathToQueryFor, {
                contoso: "http://schemas.contoso.com/review/1.0"
            });

            await context.sync();

            console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
            for (let i = 0; i < clientResult.value.length; i++) {
                console.log(clientResult.value[i]);
            }
        } else {
            console.warn("No custom XML part to query");
        }
    });
}

async function insertAttribute() {
    // Insert an attribute into a custom XML part.
    await Word.run(async (context) => {
        const settings = context.document.settings;
        const xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
        await context.sync();

        if (xmlPartIDSetting.value) {
            const customXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

            // The insertAttribute method inserts an attribute with the given name and value into the element identified by the xpath parameter.
            customXmlPart.insertAttribute("/Reviewers", { contoso: "http://schemas.contoso.com/review/1.0" }, "Nation", "US");
            const xmlBlob = customXmlPart.getXml();
            await context.sync();

            const readableXml = addLineBreaksToXML(xmlBlob.value);
            console.log("Successfully inserted attribute:");
            console.log(readableXml);
        } else {
            console.warn("No custom XML part to insert attribute into");
        }
    });
}

async function insertElement() {
    // Insert an element into a custom XML part.
    await Word.run(async (context) => {
        const settings = context.document.settings;
        const xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
        await context.sync();

        if (xmlPartIDSetting.value) {
            const customXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

            // The insertElement method inserts the given XML under the parent element identified by the xpath parameter at the provided child position index.
            customXmlPart.insertElement(
                "/Reviewers",
                "<Lead>Mark</Lead>",
                { contoso: "http://schemas.contoso.com/review/1.0" },
                0
            );
            const xmlBlob = customXmlPart.getXml();
            await context.sync();

            const readableXml = addLineBreaksToXML(xmlBlob.value);
            console.log("Successfully inserted element:");
            console.log(readableXml);
        } else {
            console.warn("No custom XML part to insert element into");
        }
    });
}

async function deleteCustomXmlPart() {
    // Delete a custom XML part.
    await Word.run(async (context) => {
        const settings = context.document.settings;
        const xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
        await context.sync();

        if (xmlPartIDSetting.value) {
            let customXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
            const xmlBlob = customXmlPart.getXml();
            customXmlPart.delete();
            customXmlPart = context.document.customXmlParts.getItemOrNullObject(xmlPartIDSetting.value);

            await context.sync();

            if (customXmlPart.isNullObject) {
                console.log(`The XML part with the ID ${xmlPartIDSetting.value} has been deleted`);

                // Delete the associated setting too.
                xmlPartIDSetting.delete();

                await context.sync();
            } else {
                const readableXml = addLineBreaksToXML(xmlBlob.value);
                const strangeMessage = `This is strange. The XML part with the id ${xmlPartIDSetting.value} wasn't deleted:\n${readableXml}`;
                console.error(strangeMessage);
            }
        } else {
            console.warn("No custom XML part to delete");
        }
    });
}

function addLineBreaksToXML(xmlBlob) {
    const replaceValue = new RegExp(">");
    return xmlBlob.replace(/></g, "> <");
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
