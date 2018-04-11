module powerbi.extensibility.visual { 
    
    import ISelectionId = powerbi.visuals.ISelectionId;
    import DataViewObject = powerbi.extensibility.utils.dataview.DataViewObject;
            
    interface ITweetListSettings {
        displayOptions: {
            showQuoted: boolean,
            showMedia: boolean;
        },
        formatOptions: {
            fontSize: number;
        }
    }    

    interface ITweet {
        image: string;
        name: string;
        screenName: string;
        date: Date;
        text: string;
        likes: number;
        retweets: number;
        sentiment: number;
        tweetJson: any;
        selectionId: ISelectionId;
        selected: boolean;
    }

    class Tweet implements ITweet{
        image: string;
        name: string;
        screenName: string;
        date: Date;
        text:string;
        likes: number;
        retweets: number;
        sentiment: number;
        tweetJson: any;
        selectionId: ISelectionId;
        selected: boolean;
    }

    export class Visual implements IVisual {
        
        private target: HTMLElement;
        private settings: ITweetListSettings;
        private host: IVisualHost;
        private selectionManager: ISelectionManager;
        private dataView: DataView;
        private selectionIds: ISelectionId[];
        private currentSelected: ISelectionId[];
        private tweets: ITweet[];
               
        constructor(options: VisualConstructorOptions) {
            this.target = options.element;
            this.host = options.host;
            this.selectionManager = options.host.createSelectionManager();      
            this.selectionIds = [];       
            this.currentSelected = [];
            this.tweets = [];
            
            this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {                       
                    this.currentSelected = ids;                    
                }
            );                           
        }

        public update(options: VisualUpdateOptions) {            
            
            this.selectionIds = this.getSelectionIds(options.dataViews[0], this.host);

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

                this.loadData();
                         
                this.createView();
            }
        }        

        private loadData(): void {            
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
            
            this.tweets = [];

            for (let index = 0; index < rows.length; index++) {
                const row: DataViewTableRow = rows[index];

                let tweet: ITweet = new Tweet();                
                tweet.selectionId = this.selectionIds[index];

                if (this.isSelectionIdInArray(this.currentSelected, tweet.selectionId)){
                    tweet.selected = true;
                }

                if (jsonColumn && row[jsonColumn] != null){                    
                    tweet.tweetJson = JSON.parse(row[jsonColumn].toString());
                }                

                if (imageColumn || tweet.tweetJson){
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.image = tweet.tweetJson.retweetedTweet.authorProfileImageUrl;
                        } else {
                            tweet.image = tweet.tweetJson.authorProfileImageUrl;
                        }                        
                    }                        
                    else if (row[imageColumn] != null){
                        tweet.image =  row[imageColumn].toString();
                    }
                }

                if (screenNameColumn || tweet.tweetJson){
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.screenName = tweet.tweetJson.retweetedTweet.screenName;
                        } else {
                            tweet.screenName = tweet.tweetJson.screenName;
                        }                        
                    }                        
                    else if (row[screenNameColumn] != null){
                        tweet.screenName =  row[screenNameColumn].toString();
                    }
                }

                if (nameColumn || tweet.tweetJson){
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.name = tweet.tweetJson.retweetedTweet.authorName;                                
                        } else {
                            tweet.name = tweet.tweetJson.authorName;
                        }
                        
                    } else if (row[nameColumn] != null) {
                        tweet.name = row[nameColumn].toString();                       
                    }
                }

                if (dateColumn || tweet.tweetJson){
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.date = tweet.tweetJson.retweetedTweet.createdDateUTC;                                
                        } else {
                            tweet.date = tweet.tweetJson.createdDateUTC;
                        }
                        
                    } else if (row[dateColumn] != null) {
                        tweet.date = new Date(row[dateColumn].toString());                       
                    }
                }

                if (textColumn || tweet.tweetJson){                    
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.text = tweet.tweetJson.retweetedTweet.text;                                
                        } else {
                            tweet.text = tweet.tweetJson.text;
                        }
                        
                    } else if (row[textColumn] != null) {
                        tweet.text = row[textColumn].toString();
                    }
                }

                if (likesColumn || tweet.tweetJson){
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.likes = tweet.tweetJson.retweetedTweet.liked;                                
                        } else {
                            tweet.text = tweet.tweetJson.liked;
                        }
                        
                    } else if (row[likesColumn] != null) {
                        tweet.likes = Number(row[likesColumn]);
                    }
                }

                if (retweetsColumn || tweet.tweetJson){
                    if (tweet.tweetJson != null){
                        if (tweet.tweetJson.retweetedTweet){
                            tweet.retweets = tweet.tweetJson.retweetedTweet.retweeted;                                
                        } else {
                            tweet.retweets = tweet.tweetJson.retweeted;
                        }
                        
                    } else if (row[textColumn] != null) {
                        tweet.retweets = Number(row[retweetsColumn]);
                    }
                }

                if (sentimentColumn){
                    if (row[sentimentColumn] != null){
                     tweet.sentiment = Number(row[sentimentColumn]);                        
                    }
                }
                
                this.tweets.push(tweet);
            }
        }
        
        private isSelectionIdInArray(selectionIds: ISelectionId[], selectionId: ISelectionId): boolean {
            if (!selectionIds || !selectionId) {
                return false;
            }

            return selectionIds.some((currentSelectionId: ISelectionId) => {
                return currentSelectionId.includes(selectionId);
            });
        }
        

        private createView(): void {                      
            while (this.target.firstChild) {
                this.target.removeChild(this.target.firstChild);
            };

            let main: HTMLDivElement = document.createElement("div");

            let isDeskTop : boolean = window.location.hostname === "pbi.microsoft.com";

            for (let i = 0; i < this.tweets.length; i++) {
                try {
                    let tweet : ITweet = this.tweets[i];                                       
                                        
                    let tweetContainer = document.createElement("div");
                    tweetContainer.classList.add("tweet-container");
                    if (tweet.selected || this.currentSelected.length == 0){
                        tweetContainer.classList.remove("dim");
                    } else if (this.currentSelected.length > 0) {
                        tweetContainer.classList.add("dim");
                    }

                    tweetContainer.onclick = (ev: MouseEvent) => {
                        this.selectionManager.select(tweet.selectionId).then((selected) => {
                            this.currentSelected = [];
                            var items = document.querySelectorAll(".tweet-container");
                            if (selected != null && selected.length > 0) {
                                for (let k = 0; k < items.length; k++) {
                                    items[k].classList.add("dim");
                                }
                                tweetContainer.classList.remove("dim");
                                this.currentSelected.push(tweet.selectionId);
                            }
                            else {
                                for (let k = 0; k < items.length; k++) {
                                    items[k].classList.remove("dim");
                                }                                
                            }
                        });
                    };
                    
                    if (tweet.tweetJson != null && tweet.tweetJson.isRetweet) {
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
                        contextText.appendChild(document.createTextNode("retweeted by " + tweet.tweetJson.authorName));
                        actionContext.appendChild(contextText);
                        
                        retweetContainer.appendChild(actionContext);
                        tweetContainer.appendChild(retweetContainer);
                    }
    
                    let tweetContentContainer = document.createElement("div");
                    tweetContentContainer.setAttribute("class", "tweet-content");
    
                    let tweetHeader = document.createElement("div");
                    tweetHeader.setAttribute("class", "tweet-header");
                    tweetHeader.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
                                            
                    if (tweet.image != null) {
                        
                        let img = document.createElement("img");
                        img.setAttribute("alt", "");
                        img.setAttribute("class", "author-photo");
                        img.setAttribute("src", tweet.image);
                            
                        tweetContentContainer.appendChild(img);
                    }
                    
                    if (tweet.name != null || tweet.screenName != null) {
    
                        if (tweet.name){
                            var nameDiv = document.createElement("div");
                            nameDiv.setAttribute("class", "author-name");
                            var strong = document.createElement("strong")
                                
                            strong.appendChild(document.createTextNode(tweet.name));
                            nameDiv.appendChild(strong);
                            tweetHeader.appendChild(nameDiv);
                        }
                        
                        let verified: boolean = false;
                        if (tweet.tweetJson != null){
                            if (tweet.tweetJson.retweetedTweet) {
                                verified = tweet.tweetJson.retweetedTweet.authorVerified;
                            } else {
                                verified = tweet.tweetJson.authorVerified;
                            }
                        }
                        
                        if (verified) {
                            let verifiedDiv = document.createElement("div");
                            verifiedDiv.setAttribute("style", `height:${this.settings.formatOptions.fontSize + 3}px;width:${this.settings.formatOptions.fontSize + 1}px;`);
                            verifiedDiv.setAttribute("class", "verified");
                            tweetHeader.appendChild(verifiedDiv);
                        }
                        if (tweet.screenName) {
                            let screenNameDiv = document.createElement("div");
                            screenNameDiv.setAttribute("class", "screen-name");
                                                                                    
                            if (tweet.screenName.length > 0 && tweet.screenName[0] !== "@"){
                                tweet.screenName = "@" + tweet.screenName;
                            } 
                            
                            screenNameDiv.appendChild(document.createTextNode(tweet.screenName));
                            tweetHeader.appendChild(screenNameDiv);
                        }
    
                        if (tweet.date){
                            let dateDiv = document.createElement("div");
                            dateDiv.setAttribute("class", "tweet-date");                                                            
                            dateDiv.appendChild(document.createTextNode("\u00A0\u2022\u00A0" + this.formatDate(tweet.date)));                        
                            
                            tweetHeader.appendChild(dateDiv);
                        }
                        tweetContentContainer.appendChild(tweetHeader);
                    }
    
                    if (tweet.tweetJson && tweet.tweetJson.replyToScreenName){
                        let replyDiv = document.createElement("div");
                        replyDiv.setAttribute("class", "in-reply-text");
                        replyDiv.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);    
                        replyDiv.appendChild(document.createTextNode("Replying to @" + tweet.tweetJson.replyToScreenName));
                        tweetContentContainer.appendChild(replyDiv);
                    }                        
                        
                    let tweetDiv = document.createElement("div");
                    tweetDiv.setAttribute("class", "tweet-text");
                    tweetDiv.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);

                    if (tweet.text){
                                                
                        tweetDiv.appendChild(document.createTextNode(tweet.text));                                                                    
                        tweetContentContainer.appendChild(tweetDiv);
                    }                                                
    
                    if (tweet.tweetJson != null && tweet.tweetJson.quotedTweet !== null) {
                                                   
                        let quoteTweetContainer = document.createElement("div");                                                
                        quoteTweetContainer.setAttribute("class", "quote-tweet-container");
    
                        if (this.settings.displayOptions.showQuoted && tweet.tweetJson.quotedTweet.media != null){
                            let quotedTweetMedia = document.createElement("div");
                            quotedTweetMedia.setAttribute("class", "quote-tweet-media");
    
                            let quotedTweetSingleMedia = document.createElement("div");
                            quotedTweetSingleMedia.setAttribute("class", "single");
    
                            let quoteTweetImg = document.createElement("img");
                            quoteTweetImg.setAttribute("alt", "");
                            quoteTweetImg.setAttribute("src", tweet.tweetJson.quotedTweet.media[0].url);
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
                        
                        strong.appendChild(document.createTextNode(tweet.tweetJson.quotedTweet.authorName));
                        nameDiv.appendChild(strong);
                        quoteAddressLine.appendChild(nameDiv);
    
                        if (tweet.tweetJson.quotedTweet.authorVerified) {
                            let verifiedDiv = document.createElement("div");
                            verifiedDiv.setAttribute("class", "verified");
                            quoteAddressLine.appendChild(verifiedDiv);
                        }
    
                        let screenNameDiv = document.createElement("div");
                        screenNameDiv.setAttribute("class", "screen-name");
                        screenNameDiv.appendChild(document.createTextNode("@" + tweet.tweetJson.quotedTweet.authorScreenName));
    
                        quoteAddressLine.appendChild(screenNameDiv);
                        quoteTweet.appendChild(quoteAddressLine);
    
                        let quotedTweetElement = document.createElement("div");
                        quotedTweetElement.setAttribute("class", "tweet-text");
                        quotedTweetElement.setAttribute("style", `font-size: ${this.settings.formatOptions.fontSize}px;`);
                        quotedTweetElement.appendChild(document.createTextNode(tweet.tweetJson.quotedTweet.text));                    
                        quoteTweet.appendChild(quotedTweetElement);
                        
                        tweetContentContainer.appendChild(tweetDiv);
    
                        quoteTweetContainer.appendChild(quoteTweet);
                        tweetContentContainer.appendChild(quoteTweetContainer);
                    }
    
                    if (this.settings.displayOptions.showMedia && tweet != null && tweet.tweetJson.media != null) {
                        
                        let mediaContainer = document.createElement("div");
                        mediaContainer.setAttribute("class", "media-container");
    
                        // ToDo: if (tweet.media.length === 1){
                        //Currently only displays first image                        
                        if (tweet.tweetJson.media.length > 0){
                            let mediaItem = tweet.tweetJson.media[0];
                            let singleContainer = document.createElement("div");
                            singleContainer.setAttribute("class", "single");
                            if (mediaItem.videoUrl === null){
                                let img = document.createElement("img");
                                img.setAttribute("alt", "");
                                img.setAttribute("src", tweet.tweetJson.media[0].url);
                                singleContainer.appendChild(img);                                
                            } else {                                     
                                if (!isDeskTop){
                                    let video = document.createElement("video");  
                                    let duration: number = null;                            
                                    let currentTime: number = null;
                                    video.setAttribute("class", "video-player")                                                              
                                    video.setAttribute("poster", mediaItem.url);
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
                                        ev.preventDefault();
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
                                    source.setAttribute("src", mediaItem.videoUrl);
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
                                        img.setAttribute("src", mediaItem.url);
                                        singleContainer.appendChild(img);                                        
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
                    
                    if (tweet.retweets) {
                        let divBg = document.createElement("div");
                        divBg.setAttribute("class", "retweet");
                        divBg.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;width:${this.settings.formatOptions.fontSize+2}px;`);
                        footerDiv.appendChild(divBg);
    
                        let div = document.createElement("div");
                        div.setAttribute("class", "footer-value");
                        div.appendChild(document.createTextNode(this.formatNumber(tweet.retweets)));
                        footerDiv.appendChild(div);
                    }
    
                    if (tweet.likes) {
                        let divBg = document.createElement("div");
                        divBg.setAttribute("class", "like");
                        divBg.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;width:${this.settings.formatOptions.fontSize + 2}px;`);
                        footerDiv.appendChild(divBg);
    
                        let div = document.createElement("div");
                        div.setAttribute("class", "footer-value");    
                        div.appendChild(document.createTextNode(this.formatNumber(tweet.likes)));                                            
                        footerDiv.appendChild(div);
                    }
                        
                    if (tweet.sentiment){
                        let divBg = document.createElement("div");
                        divBg.setAttribute("class", "sentiment");
                        divBg.setAttribute("style", `height:${this.settings.formatOptions.fontSize}px;width:${this.settings.formatOptions.fontSize + 2}px;`);
                        footerDiv.appendChild(divBg);
    
                        let div = document.createElement("div");
                        div.setAttribute("class", "footer-value");    
                                               
                        if (tweet.sentiment < 30){
                            divBg.classList.add("negative");
                        } else if (tweet.sentiment > 70){
                            divBg.classList.add("positive");
                        } else {
                            divBg.classList.add("neutral");
                        }
                        div.appendChild(document.createTextNode(tweet.sentiment.toFixed(0)));
                                            
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
            let value: Fill;
            value = DataViewObject.getValue(objects, propertyName);
            if (!value || !value.solid) {
                return defaultColor;
            }

            return value.solid.color; 
        }

        private getColumnIndex(columns: DataViewMetadataColumn[], columnType: string): number {
            if (columns) {
                for (let i = 0; i < columns.length; i++) {
                    let column = columns[i];
                    if (column.objects){
                        let object = column.objects["columnOptions"];
                        if (object) {
                            let property: string = object["fieldType"].toString();
                            if (property === columnType)
                                return i;
                        }
                    }
                }
            }
            return null;
        }

        private getDataColumnIndexes(columns: DataViewMetadataColumn[], columnType: string): number[] {
            if (columns) {
                let indexes: number[] = [];
                for (let i = 0; i < columns.length; i++) {
                    let column = columns[i];
                    if (column.objects){
                        let object = column.objects["columnOptions"];
                        if (object) {
                            let property: string = object["fieldType"].toString();
                            if (property === columnType)
                                indexes.push(i);
                        }
                    }                    
                }
                return indexes;
            }
            return null;
        }
    }
}