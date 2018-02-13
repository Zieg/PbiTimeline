module powerbi.extensibility.visual {
    "use strict";
            
    interface ITweetListSettings {
        displayOptions: {
            showQuoted: boolean,
            showMedia: boolean;
        },
        formatOptions: {
            fontSize: number;
        }
    }

    export class Visual implements IVisual {
        
        private target: HTMLElement;
        private settings: ITweetListSettings;
        private host: IVisualHost;
        private selectionManager: ISelectionManager;
        private dataView: DataView;        
               
        constructor(options: VisualConstructorOptions) {
            this.target = options.element;
            this.host = options.host;
            this.selectionManager = options.host.createSelectionManager();              
        }

        private escapeRegExp(str : string) {
            return str.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
        }

        private replaceAll(str : string, find : string, replace : string) {
            return str.replace(new RegExp(this.escapeRegExp(find), 'g'), replace);
        }
        
        public update(options: VisualUpdateOptions) {
            
            while (this.target.firstChild) {
                this.target.removeChild(this.target.firstChild);
            };

            let ids: ISelectionId[] = [];

            try{
                ids = this.getSelectionIds(options.dataViews[0], this.host);
            }
            catch (e){
                
            }                            

            var dataView: DataView = this.dataView = options.dataViews[0];

            if (dataView != null && dataView.table != null) {
                let table: DataViewTable = dataView.table;
                this.target.style.overflowY = "auto";

                let columns: DataViewMetadataColumn[] = table.columns;
                
                let defaultSettings: ITweetListSettings = {
                    displayOptions: {
                        showQuoted: true,
                        showMedia: false
                    },
                    formatOptions: {
                        fontSize: 12
                    }
                };

                let objects = options.dataViews[0].metadata.objects;

                let currentSettings: ITweetListSettings = {
                    displayOptions: {
                        showMedia: this.getValue<boolean>(objects, "tweetOptions", "showMedia", defaultSettings.displayOptions.showMedia),
                        showQuoted: this.getValue<boolean>(objects, "tweetOptions", "showQuoted", defaultSettings.displayOptions.showQuoted)
                    },
                    formatOptions: {
                        fontSize: this.getValue<number>(objects, "formatting", "fontSize", defaultSettings.formatOptions.fontSize)
                    }
                }

                this.settings = currentSettings;
                         
                this.createView(ids);
            }
        }

        private createView(selectionIds: ISelectionId[]): void {            
            let main: HTMLDivElement = document.createElement("div");

            let table: DataViewTable = this.dataView.table;
            let columns: DataViewMetadataColumn[] = table.columns;
            let rows = table.rows;

            let imageColumn = this.getColumnIndex(columns, "authorProfilePicture");
            let screenNameColumn = this.getColumnIndex(columns, "authorScreenName");
            let nameColumn = this.getColumnIndex(columns, "authorName");
            let dateColumn = this.getColumnIndex(columns, "tweetDate");
            let jsonColumn = this.getColumnIndex(columns, "tweetJson");
            let textColumn = this.getColumnIndex(columns, "tweetText");

            let likesColumn = this.getColumnIndex(columns, "likes");
            let retweetsColumn = this.getColumnIndex(columns, "retweets");

            let sentimentColumn = this.getColumnIndex(columns, "sentiment");

            let isDeskTop : boolean = window.location.hostname === "pbi.microsoft.com";

            for (let i = 0; i < rows.length; i++) {


                try{
                    let row : DataViewTableRow = rows[i];
                    let selectionId : ISelectionId = selectionIds[i];
                    
                    let tweet : any = null;
                    if (jsonColumn != null){                    
                        let json: string = row[jsonColumn].toString();
                        tweet = JSON.parse(json);
                    }
                    
                    let tweetContainer = document.createElement("div");
                    tweetContainer.setAttribute("class", "tweet-container");   
                    
                    if (tweet != null && tweet.isRetweet) {
                        let retweetContainer = document.createElement("div");
                        retweetContainer.setAttribute("class", "tweet-actions");   
    
                        let actionContext = document.createElement("div");
                        actionContext.setAttribute("class", "action-context");
    
                        let retweetIcon = document.createElement("div");
                        retweetIcon.setAttribute("class", "retweet-image");
                        retweetIcon.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;`);                    
                        actionContext.appendChild(retweetIcon);
    
                        let contextText = document.createElement("div");
                        contextText.setAttribute("class", "context-text");
                        contextText.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
                        contextText.appendChild(document.createTextNode("retweeted by " + this.htmlEscape(tweet.authorName)));
                        actionContext.appendChild(contextText);
                        
                        retweetContainer.appendChild(actionContext);
                        tweetContainer.appendChild(retweetContainer);
                    }
    
                    let tweetContentContainer = document.createElement("div");
                    tweetContentContainer.setAttribute("class", "tweet-content");
    
                    let tweetHeader = document.createElement("div");
                    tweetHeader.setAttribute("class", "tweet-header");
                    tweetHeader.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
    
                    this.setclick(tweetHeader, selectionId, tweetContainer);
                    
                    if (tweet != null || imageColumn != null) {
                        
                        let img = document.createElement("img");
                        img.setAttribute("alt", "");
                        img.setAttribute("class", "author-photo");
                        if (tweet != null){
                            if (tweet.retweetedTweet){
                                img.setAttribute("src", encodeURI(tweet.retweetedTweet.authorProfileImageUrl));
                            } else {
                                img.setAttribute("src", encodeURI(tweet.authorProfileImageUrl));
                            }                        
                        }                        
                        else 
                            img.setAttribute("src", encodeURI(row[imageColumn].toString()));
                        tweetContentContainer.appendChild(img);
                    }
                    
                    if (tweet != null || nameColumn != null || screenNameColumn != null) {
    
                        if (nameColumn != null || tweet != null){
                            var nameDiv = document.createElement("div");
                            nameDiv.setAttribute("class", "author-name");
                            var strong = document.createElement("strong")
    
                            let name: string = "";
    
                            if (tweet != null){
                                if (tweet.retweetedTweet){
                                    name = tweet.retweetedTweet.authorName;                                
                                } else {
                                    name = tweet.authorName;
                                }
                                
                            } else {
                                if (row[nameColumn] != null){
                                    name = row[nameColumn].toString();
                                }                        
                            }
                            strong.appendChild(document.createTextNode(this.htmlEscape(name)));
                            nameDiv.appendChild(strong);
                            tweetHeader.appendChild(nameDiv);
                        }
                        
                        let verified: boolean = false;
                        if (tweet != null){
                            if (tweet.retweetedTweet) {
                                verified = tweet.retweetedTweet.authorVerified;
                            } else {
                                verified = tweet.authorVerified;
                            }
                        }
                        
                        if (verified) {
                            let verifiedDiv = document.createElement("div");
                            verifiedDiv.setAttribute("style", `height:${this.settings.formatOptions.fontSize + 3}px;width:${this.settings.formatOptions.fontSize + 1}px;`);
                            verifiedDiv.setAttribute("class", "verified");
                            tweetHeader.appendChild(verifiedDiv);
                        }
                        if (tweet != null || screenNameColumn != null) {
                            let screenNameDiv = document.createElement("div");
                            if (tweet != null || nameColumn != null){
                                screenNameDiv.setAttribute("class", "screen-name");
                            }
                            
                            let screenNameVal : string = "";
                            if (tweet != null){
                                if (tweet.retweetedTweet){
                                    screenNameVal = tweet.retweetedTweet.authorScreenName;
                                } else {
                                    screenNameVal = tweet.authorScreenName;
                                }
                                
                            } else {
                                if (row[screenNameColumn] != null)
                                    screenNameVal = row[screenNameColumn].toString();
                            }
                            
                            if (screenNameVal.length > 0 && screenNameVal[0] !== "@"){
                                screenNameVal = "@" + screenNameVal;
                            } 
                            
                            screenNameDiv.appendChild(document.createTextNode(this.htmlEscape(screenNameVal)));
                            tweetHeader.appendChild(screenNameDiv);
                        }
    
                        if (tweet != null || dateColumn != null){
                            let dateDiv = document.createElement("div");
                            dateDiv.setAttribute("class", "tweet-date");
    
                            let tweetDate: Date = new Date();
    
                            if (tweet != null){
                                if (tweet.retweetedTweet){                                                
                                    tweetDate = new Date(tweet.retweetedTweet.createdDateUTC);
                                } else {
                                    tweetDate = new Date(tweet.createdDateUTC);
                                } 
                            } else {
                                if (row[dateColumn] != null)
                                    tweetDate = new Date(row[dateColumn].toString());
                            }
                            
                            dateDiv.appendChild(document.createTextNode("\u00A0\u2022\u00A0" + this.formatDate(tweetDate)));                        
                            
                            tweetHeader.appendChild(dateDiv);
                        }
                        tweetContentContainer.appendChild(tweetHeader);
                    }
    
                    if (tweet != null && tweet.replyToScreenName !== null){                                                            
                        let replyDiv = document.createElement("div");
                        replyDiv.setAttribute("class", "in-reply-text");
                        replyDiv.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);    
                        replyDiv.appendChild(document.createTextNode("Replying to @" + this.htmlEscape(tweet.replyToScreenName)));
                        tweetContentContainer.appendChild(replyDiv);
                    }
    
                    let tweetDiv = document.createElement("div");
                    tweetDiv.setAttribute("class", "tweet-text");
                    tweetDiv.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
    
                    let tweetText: string = "";
    
                    if (tweet != null){
                        if (tweet.retweetedTweet){                    
                            tweetText = tweet.retweetedTweet.text;
                        } else {
                            tweetText = tweet.text;
                        }  
                    } else if (textColumn != null){
                        tweetText = row[textColumn].toString();
                    }
    
                    tweetDiv.appendChild(document.createTextNode(this.htmlEscape(tweetText)));                
                                    
                    this.setclick(tweetDiv, selectionId, tweetContainer);
    
                    tweetContentContainer.appendChild(tweetDiv);
    
                    if (tweet != null && tweet.quotedTweet !== null) {
                        
                        debugger;
    
                        let quoteTweetContainer = document.createElement("div");
    
                        this.setclick(quoteTweetContainer, selectionId, tweetContainer);
                        
                        quoteTweetContainer.setAttribute("class", "quote-tweet-container");
    
                        if (this.settings.displayOptions.showQuoted && tweet.quotedTweet.media != null){
                            let quotedTweetMedia = document.createElement("div");
                            quotedTweetMedia.setAttribute("class", "quote-tweet-media");
    
                            let quotedTweetSingleMedia = document.createElement("div");
                            quotedTweetSingleMedia.setAttribute("class", "single");
    
                            let quoteTweetImg = document.createElement("img");
                            quoteTweetImg.setAttribute("alt", "");
                            quoteTweetImg.setAttribute("src", encodeURI(tweet.quotedTweet.media[0].url));
                            quoteTweetImg.setAttribute("style", "height:100%;");
                            quotedTweetSingleMedia.appendChild(quoteTweetImg);
    
                            quotedTweetMedia.appendChild(quotedTweetSingleMedia);
                            quoteTweetContainer.appendChild(quotedTweetMedia);
                        }
    
                        let quoteTweet = document.createElement("div");
                        quoteTweet.setAttribute("class", "quote-tweet");
                        quoteTweet.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
    
                        let quoteAddressLine = document.createElement("div");
                        quoteAddressLine.setAttribute("class", "tweet-header");
    
                        var nameDiv = document.createElement("div");
                        nameDiv.setAttribute("class", "author-name");
                        var strong = document.createElement("strong");
                        
                        strong.appendChild(document.createTextNode(this.htmlEscape(tweet.quotedTweet.authorName)))
                        nameDiv.appendChild(strong);
                        quoteAddressLine.appendChild(nameDiv);
    
                        if (tweet.quotedTweet.authorVerified) {
                            let verifiedDiv = document.createElement("div");
                            verifiedDiv.setAttribute("class", "verified");
                            quoteAddressLine.appendChild(verifiedDiv);
                        }
    
                        let screenNameDiv = document.createElement("div");
                        screenNameDiv.setAttribute("class", "screen-name");
                        screenNameDiv.appendChild(document.createTextNode("@" + this.htmlEscape(tweet.quotedTweet.authorScreenName)));
    
                        quoteAddressLine.appendChild(screenNameDiv);
                        quoteTweet.appendChild(quoteAddressLine);
    
                        let quotedTweetElement = document.createElement("div");
                        quotedTweetElement.setAttribute("class", "tweet-text");
                        quotedTweetElement.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
                        quotedTweetElement.appendChild(document.createTextNode(tweet.quotedTweet.text));                    
                        quoteTweet.appendChild(quotedTweetElement);
                        
                        tweetContentContainer.appendChild(tweetDiv);
    
                        quoteTweetContainer.appendChild(quoteTweet);
                        tweetContentContainer.appendChild(quoteTweetContainer);
                    }
    
                    if (this.settings.displayOptions.showMedia && tweet != null && tweet.media != null) {
                        
                        let mediaContainer = document.createElement("div");
                        mediaContainer.setAttribute("class", "media-container");
    
                        // ToDo: if (tweet.media.length === 1){
                        //Currently only displays first image                        
                        if (tweet.media.length > 0){
                            let mediaItem = tweet.media[0];
                            let singleContainer = document.createElement("div");
                            singleContainer.setAttribute("class", "single");
                            if (mediaItem.videoUrl === null){
                                let img = document.createElement("img");
                                img.setAttribute("alt", "");
                                img.setAttribute("src", encodeURI(tweet.media[0].url));
                                singleContainer.appendChild(img);
                                this.setclick(singleContainer, selectionId, tweetContainer);
                            } else {                                     
                                if (!isDeskTop){
                                    let video = document.createElement("video");  
                                    let duration: number = null;                            
                                    let currentTime: number = null;
                                    video.setAttribute("class", "video-player")                                                              
                                    video.setAttribute("poster", encodeURI(mediaItem.url));
                                    video.setAttribute("preload", "metadata");
        
                                    let progress = document.createElement("progress")
                                    progress.value = 0;
        
                                    let videoInfo = document.createElement("div");
                                    videoInfo.setAttribute("class", "info play");
                                    videoInfo.innerText = "loading..."
        
                                    let controls = document.createElement("div");
                                    controls.setAttribute("class", "controls hide");
                                    
                                    singleContainer.onclick = (ev: MouseEvent) => {                                
                                        if (video.paused) {
                                            video.play();
                                            videoInfo.classList.add("pause");
                                            videoInfo.classList.remove("play");
                                        } else {
                                            video.pause();
                                            videoInfo.classList.remove("pause");
                                            videoInfo.classList.add("play");                   
                                        }
                                    }; 
        
                                    video.onloadedmetadata = (ev: Event) => {
                                        duration = video.duration;
                                        videoInfo.innerText = this.toFormattedTime(duration);
                                        progress.max = video.duration;                                
                                    };
        
                                    video.ontimeupdate = (ev: Event) => {                                
                                        currentTime = video.currentTime;                                  
                                        videoInfo.innerText = this.toFormattedTime(currentTime) + " / " + this.toFormattedTime(duration);
                                        progress.value = video.currentTime;
                                    }
        
                                    singleContainer.onmouseenter = (ev:MouseEvent) => {
                                        controls.classList.remove("hide");                                
        
                                    };
                                    singleContainer.onmouseleave = (ev: MouseEvent)=> {
                                        controls.classList.add("hide");
                                    };                            
        
                                    let source = document.createElement("source");
                                    source.setAttribute("src", encodeURI(mediaItem.videoUrl));
                                    source.setAttribute("type", "video/mp4");
                                    video.appendChild(source);      
                                                               
                                    singleContainer.appendChild(video);       
                                    controls.appendChild(progress);
                                    controls.appendChild(videoInfo);
                                    singleContainer.appendChild(controls); 
                                } else {
                                    if (mediaItem.url){
                                        let img = document.createElement("img");                                    
                                        img.setAttribute("alt", "");
                                        img.setAttribute("src", encodeURI(mediaItem.url));
                                        singleContainer.appendChild(img);
                                        this.setclick(singleContainer, selectionId, tweetContainer);
                                    }                                
                                }                                                 
                            }
                            
                            mediaContainer.appendChild(singleContainer);
                        }
    
                        tweetContentContainer.appendChild(mediaContainer);
                        
                    }
    
                    let footerDiv = document.createElement("div");
                    footerDiv.setAttribute("class", "item-footer");
                    footerDiv.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
    
                    this.setclick(footerDiv, selectionId, tweetContainer);
                    
                    if (tweet != null || retweetsColumn != null) {
                        let divBg = document.createElement("div");
                        divBg.setAttribute("class", "retweet");
                        divBg.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;width:${this.settings.formatOptions.fontSize+2}px;`);
                        footerDiv.appendChild(divBg);
    
                        let div = document.createElement("div");
                        div.setAttribute("class", "footer-value");
                        if (retweetsColumn != null) {
                            if (row[retweetsColumn] != null){
                                div.appendChild(document.createTextNode(this.formatNumber(parseInt(row[retweetsColumn].toString()))));
                            }  else {
                                div.appendChild(document.createTextNode("0"));
                            }                     
                        } else {
                            div.appendChild(document.createTextNode(this.formatNumber(tweet.retweeted)));
                        }
                        footerDiv.appendChild(div);
                    }
    
                    if (tweet != null || likesColumn != null) {
                        let divBg = document.createElement("div");
                        divBg.setAttribute("class", "like");
                        divBg.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;width:${this.settings.formatOptions.fontSize + 2}px;`);
                        footerDiv.appendChild(divBg);
    
                        let div = document.createElement("div");
                        div.setAttribute("class", "footer-value");    
                        if (likesColumn != null) {
                            if (row[likesColumn] != null){
                                div.appendChild(document.createTextNode(this.formatNumber(parseInt(row[likesColumn].toString()))));
                            } else {
                                div.appendChild(document.createTextNode("0"));
                            }                        
                        } else {
                            div.appendChild(document.createTextNode(this.formatNumber(tweet.liked)));
                        }
                                            
                        footerDiv.appendChild(div);
                    }
    
                    if (sentimentColumn != null){
                        let divBg = document.createElement("div");
                        divBg.setAttribute("class", "sentiment");
                        divBg.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;width:${this.settings.formatOptions.fontSize + 2}px;`);
                        footerDiv.appendChild(divBg);
    
                        let div = document.createElement("div");
                        div.setAttribute("class", "footer-value");    
                        if (sentimentColumn != null) {
                            if (row[sentimentColumn] != null){
                                let number = parseFloat(row[sentimentColumn].toString());                            
                                if (number < 0.4){
                                    divBg.classList.add("negative");
                                } else if (number > 0.6){
                                    divBg.classList.add("positive");
                                } else {
                                    divBg.classList.add("neutral");
                                }
                                div.appendChild(document.createTextNode(number.toFixed(2)));
                            } else {
                                div.appendChild(document.createTextNode(""));
                            }                        
                        }
                                            
                        footerDiv.appendChild(div);
                    }
    
                    tweetContentContainer.appendChild(footerDiv);
                    tweetContainer.appendChild(tweetContentContainer);    
                    main.appendChild(tweetContainer); 
                }
                catch(e){
                    console.warn(e);
                }

                               
            }     
                        
            this.target.appendChild(main);
        }

        private htmlEscape(val: string) : string {
            return val
                .replace(/&/g, '&amp;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
        }

        private toFormattedTime(time: number) : string {
            let minutes = Math.floor(time / 60);
            let seconds = Math.floor(time % 60);     
            return minutes.toFixed(0) + ":" + this.formatTime(seconds);
        }

        private formatTime(value: number) : string{
            if (value < 10){
                return "0" + value.toFixed(0);
            }
            return value.toFixed(0);
        }
               
        private setclick(htmlDiv: HTMLDivElement, selectionId: ISelectionId, tweetContainer: HTMLDivElement) {
            htmlDiv.onclick = (ev: MouseEvent) => {
                this.selectionManager.select(selectionId).then((selected) => {
                    var items = document.querySelectorAll(".tweet-container");
                    if (selected != null && selected.length > 0) {
                        for (let k = 0; k < items.length; k++) {
                            items[k].classList.add("dim");
                        }
                        tweetContainer.classList.remove("dim");
                    }
                    else {
                        for (let k = 0; k < items.length; k++) {
                            items[k].classList.remove("dim");
                        }
                    }
                });
            };
        }

        private formatDate(date: Date) : string {            
            let one_day=1000*60*60*24;
            let one_hour=1000*60*60;
            let one_minute=1000*60;

            let difference = Date.now() - date.getTime();
            if (difference < one_hour){
                let minutes = Math.round(difference / one_minute);
                return minutes + "m";
            }
            else if (difference < one_day){
                let hours = Math.round(difference / one_hour);
                return hours + "h";
            } 
            else {
                var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                let returnVal = monthNames[date.getMonth()] + " " + date.getDate();

                if (date.getFullYear() !== new Date().getFullYear())
                    returnVal += " " + date.getFullYear(); 

                return returnVal;
            }            
        }
        
        private formatNumber(number: number): string {            
            if (number < 10000)
                return number.toLocaleString();
            if (number < 1000000) {
                const result = (number / 1000).toFixed(1);
                if (result.charAt(result.length - 1) === "0")
                    return (number / 1000).toFixed(0) + "K";
                return result + "K";
            }
            if (number < 1000000000) {
                const result = (number / 1000000).toFixed(1);
                if (result.charAt(result.length - 1) === "0")
                    return (number / 1000000).toFixed(0) + "M";
                return result + "M";
            } else {
                const result = (number / 1000000000).toFixed(1);
                if (result.charAt(result.length - 1) === "0")
                    return (number / 1000000000).toFixed(0) + "B";
                return result + "B";
            }
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {

            let objectName = options.objectName;
            let objectEnumeration: VisualObjectInstance[] = [];
            var columns: DataViewMetadataColumn[] = this.dataView.table.columns;        

            switch (objectName) {
                case 'tweetOptions':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            showMedia: this.settings.displayOptions.showMedia,
                            showQuoted: this.settings.displayOptions.showQuoted
                        },
                        selector: null
                    });
                    break;
                case 'columnOptions':
                    for (var i = 0; i < columns.length; i++) {
                        var currentColumn: DataViewMetadataColumn = columns[i];
                        objectEnumeration.push({
                            objectName: objectName,
                            displayName: currentColumn.displayName,
                            properties: {
                                fieldType: this.getValue<string>(currentColumn.objects, objectName, "fieldType", "notMapped")
                            },
                            selector: { metadata: currentColumn.queryName }
                        });
                    };
                    break;
                case 'formatting':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            fontSize: this.settings.formatOptions.fontSize
                        },
                        validValues: {
                            fontSize: {
                                numberRange: {
                                    min: 6,
                                    max: 24
                                }
                            }
                        },
                        selector: null
                    });
                    break;                
            }

            return objectEnumeration;
        }

        /**
         * Ids for all rows in the table as proposed here:
         * https://github.com/Microsoft/PowerBI-visuals/issues/77
         * https://github.com/Microsoft/PowerBI-visuals/issues/248
         * 
         * @param dataView 
         * @param host 
         */
        private getSelectionIds(dataView: DataView, host: IVisualHost): ISelectionId[] {
            return dataView.table.identity.map((identity: DataViewScopeIdentity) => {
                const categoryColumn: DataViewCategoryColumn = {
                    source: dataView.table.columns[0],
                    values: null,
                    identity: [identity]
                };

                return host.createSelectionIdBuilder()
                    .withCategory(categoryColumn, 0)
                    .createSelectionId();
            });
        }

        private getValue<T>(objects: DataViewObjects, objectName: string, propertyName: string, defaultValue: T): T {
            if (objects) {
                let object = objects[objectName];
                if (object) {
                    let property: T = <T>object[propertyName];
                    if (property !== undefined) {
                        return property;
                    }
                }
            }
            return defaultValue;
        }

        private getValueFromObject<T>(object: DataViewObject, objectName: string, propertyName: string, defaultValue: T): T {
            let property: T = <T>object[propertyName];
            if (property !== undefined) {
                return property;
            }
            return defaultValue;
        }

        private getFillColorByPropertyName(objects: DataViewObjects, objectName: string, propertyName: string, defaultColor?: string): string {
            if (objects) {
                let object = objects[objectName];
                if (object) {
                    let value: Fill = this.getValueFromObject(object, objectName, propertyName, defaultColor);
                    if (!value || !value.solid) {
                        return defaultColor;
                    }

                    return value.solid.color;
                }
            }
            return defaultColor;
        }

        private getColumnIndex(columns: DataViewMetadataColumn[], columnType: string): number {
            try{
                if (columns) {
                    for (let i = 0; i < columns.length; i++) {
                        let column = columns[i];
                        let object = column.objects["columnOptions"];
                        if (object) {
                            let property: string = object["fieldType"].toString();
                            if (property === columnType)
                                return i;
                        }
                    }    
                }
                return null;
            }
            catch (e){                
                return null;
            }
            
        }

        private getDataColumnIndexes(columns: DataViewMetadataColumn[], columnType: string): number[] {
            if (columns) {
                let indexes: number[] = [];
                for (let i = 0; i < columns.length; i++) {
                    let column = columns[i];
                    let object = column.objects["columnOptions"];
                    if (object) {
                        let property: string = object["fieldType"].toString();
                        if (property === columnType)
                            indexes.push(i);
                    }
                }
                return indexes;
            }
            return null;
        }
    }
}