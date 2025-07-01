#targetengine "session"
#target indesign

function countCharactersInStoryChains() {
    var doc = app.activeDocument;
    if (!doc) {
        alert("Нет открытого документа!");
        return;
    }

    var targetLayer = doc.activeLayer;
    if (!targetLayer) {
        alert("Не удалось определить активный слой!");
        return;
    }

    var resultsLayer = doc.layers.item("Results");
    if (!resultsLayer.isValid) {
        resultsLayer = doc.layers.add({name: "Results"});
    }

    var stories = doc.stories;
    var processedStories = {};

    // Получаем все мастер-страницы документа
    var masterSpreads = doc.masterSpreads;
    var masterPages = [];
    for (var m = 0; m < masterSpreads.length; m++) {
        masterPages = masterPages.concat(masterSpreads[m].pages);
    }
    
    for (var i = 0; i < stories.length; i++) {
        var story = stories[i];
        
        if (processedStories[story.id]) continue;
        
        try {
            var firstFrame = story.textContainers[0];
            if (!firstFrame || firstFrame.itemLayer.name !== targetLayer.name) continue;
            
            // НАДЕЖНАЯ ПРОВЕРКА НА ШАБЛОННЫЕ СТРАНИЦЫ
            var isOnMasterPage = false;
            var parentPage = firstFrame.parentPage;
            
            if (parentPage) {
                // Проверяем, есть ли эта страница в списке мастер-страниц
                for (var p = 0; p < masterPages.length; p++) {
                    if (masterSpreads[p].name.split("-")[0] === parentPage.name) {
                        isOnMasterPage = true;
                        break;
                    }
                }
            }
            
            if (isOnMasterPage) continue;
            
            // Подсчет символов
            var textContent = story.contents.replace(/\s+$/, '');
            textContent = textContent.replace(/[\r\n\u2029\u0003\u0005]/g, '');
            var charCount = textContent.length;
            
            var topFrame = findTopmostFrameInStory(story);
            if (!topFrame) continue;
            
            createResultFrame(topFrame, charCount, resultsLayer);
            processedStories[story.id] = true;
            
        } catch (e) {
            $.writeln("Ошибка при обработке истории: " + e.message);
        }
    }
    
    alert("Подсчет завершен! Результаты на слое 'Results'.\nШаблонные страницы исключены.");
    
    function findTopmostFrameInStory(story) {
        var frames = story.textContainers;
        if (frames.length === 0) return null;
        
        var topFrame = frames[0];
        var minY = topFrame.geometricBounds[0];
        
        for (var j = 1; j < frames.length; j++) {
            try {
                var currentY = frames[j].geometricBounds[0];
                if (currentY < minY) {
                    minY = currentY;
                    topFrame = frames[j];
                }
            } catch (e) {
                $.writeln("Ошибка при анализе фрейма: " + e.message);
            }
        }
        return topFrame;
    }
    
    function createResultFrame(targetFrame, count, layer) {
        try {
            var parent = targetFrame.parent;
            var page = (parent.constructor.name === "Page") ? parent : targetFrame.parentPage;
            
            if (!page) {
                $.writeln("Не удалось определить страницу для фрейма");
                return;
            }
            
            var resultFrame = page.textFrames.add({
                geometricBounds: [
                    targetFrame.geometricBounds[0],
                    targetFrame.geometricBounds[1],
                    targetFrame.geometricBounds[0] + 20,
                    targetFrame.geometricBounds[1] + 30
                ],
                contents: count.toLocaleString(),
                itemLayer: layer
            });
                                   
            
            var text = resultFrame.texts[0];
            text.paragraphs[0].justification = Justification.CENTER_ALIGN;
            if (text) {
                text.paragraphs[0].pointSize = 20;
                text.paragraphs[0].fillColor = doc.swatches.item("Black");
                resultFrame.textFramePreferences.firstBaselineOffset = FirstBaseline.LEADING_OFFSET;
                resultFrame.textFramePreferences.insetSpacing = [0, 0, 0, 0]; // [Top, Left, Bottom, Right]
                resultFrame.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
                resultFrame.textFramePreferences.ignoreWrap = true;
                text.paragraphs[0].firstLineIndent = 0;     // Абзацный отступ = 0
                text.paragraphs[0].spaceBefore = 0;         // Отступ перед абзацем = 0
                text.paragraphs[0].spaceAfter = 0;          // Отступ после абзаца = 0
            
                // Междустрочный интервал (leading) = 15 pt
                text.paragraphs[0].leading = 15;
            }

            // resultFrame.fillColor = doc.swatches.item("Paper");
            try {
                // Создаем или находим цвет CMYK 0/0/90/10
                var yellowColor = doc.swatches.item("Custom_Yellow");
                if (!yellowColor.isValid) {
                    yellowColor = doc.swatches.add();
                    yellowColor.model = ColorModel.PROCESS;
                    yellowColor.colorValue = [0, 0, 90, 10]; // C=0, M=0, Y=90, K=10
                    yellowColor.name = "Custom_Yellow";
                }
                resultFrame.fillColor = yellowColor;
            } catch(e) {
                $.writeln("Ошибка при создании цвета: " + e.message);
                resultFrame.fillColor = doc.swatches.item("Paper"); // fallback
            }

            resultFrame.strokeWeight = 0.5;
            resultFrame.strokeColor = doc.swatches.item("Black");
            
        } catch (e) {
            $.writeln("Ошибка при создании фрейма: " + e.message);
        }
    }
}

countCharactersInStoryChains();
