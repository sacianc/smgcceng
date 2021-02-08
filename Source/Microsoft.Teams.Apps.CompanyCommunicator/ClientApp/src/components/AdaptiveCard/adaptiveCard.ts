export const getInitAdaptiveCard = () => {
    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": "Title",
                    "size": "ExtraLarge",
                    "wrap": true
                },
                {
                    "type": "Image",
                    "spacing": "Default",
                    "url": "",
                    "size": "Stretch",
                    "width": "400px",
                    "altText": ""
                },
                {
                    "type": "TextBlock",
                    "text": "",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0"
        }
    );
}

export const getCardTitle = (card: any) => {
    return card.body[0].text;
}

export const setCardTitle = (card: any, title: string) => {
    card.body[0].text = title;
}

export const getCardImageLink = (card: any) => {
    return card.body[1].url;
}

export const setCardImageLink = (card: any, imageLink?: string) => {
    card.body[1].url = imageLink;
}

export const getCardSummary = (card: any) => {
    return card.body[2].text;
}

export const setCardSummary = (card: any, summary?: string) => {
    card.body[2].text = summary;
}

export const getCardAuthor = (card: any) => {
    return card.body[3].text;
}

export const setCardAuthor = (card: any, author?: string) => {
    card.body[3].text = author;
}

export const getCardBtnTitle = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLink = (card: any) => {
    return card.actions[0].url;
}

export const setCardBtn = (card: any, buttonTitle?: string, buttonLink?: string, buttonTitle2?: string, buttonLink2?: string) => {
    if ((buttonTitle && buttonLink) && (buttonTitle2 && buttonLink2)) {
        card.actions = [
            {
                "type": "Action.OpenUrl",
                "title": buttonTitle,
                "url": buttonLink
            },
            {
                "type": "Action.OpenUrl",
                "title": buttonTitle2,
                "url": buttonLink2
            }
        ];
    }
    else if ((buttonTitle && buttonLink) && (!buttonTitle2 || !buttonLink2)) {
        card.actions = [
            {
                "type": "Action.OpenUrl",
                "title": buttonTitle,
                "url": buttonLink
            }
        ];
    }
    else if ((!buttonTitle || !buttonLink) && (buttonTitle2 && buttonLink2)) {
        card.actions = [
            {
                "type": "Action.OpenUrl",
                "title": buttonTitle2,
                "url": buttonLink2
            }
        ];
    }
    else {
        delete card.actions;
    }
}
