```mermaid
graph TD
    A[AWS S3] -->|Download 'Review Cases.xlsx'| B[Excel]
    B -->|Extract the case number and <br>send to Salesforce| C(Salesforce)
    C -->|Find the Case Contact Name's <br>Mailing Address and look up <br>property address| D[www.google.com/maps]
    D -->|Add maps screenshot as a FeedItem| C
    C -->|Once case FeedItem has been updated <br>send message to Teams channel| E[Teams]
```

# Bot Description

The bot starts off by connecting to AWS S3 to download an excel file, extract the data from that excel file in the form of Salesforce case number, and then uses those case numbers to extract contact addresses. Those addresses are then entered into Google Maps and a screenshot is taken and saved to an output folder. Once completed it uploads the images as a FeedItem to the Salesforce case and finally sends a teams message alearting the channel that the case has been updated.

# Template: Standard Robot Framework

Want to get started using [Robot Framework](https://robocorp.com/docs/languages-and-frameworks/robot-framework/basics) this is the simplest template to start from.

This template robot:

- Uses [Robot Framework](https://robocorp.com/docs/languages-and-frameworks/robot-framework/basics) syntax.
- Includes all the necessary dependencies and initialization commands (`conda.yaml`).
- Provides a simple task template to start from (`tasks.robot`).

## Learning materials

- [All docs related to Robot Framework](https://robocorp.com/docs/languages-and-frameworks/robot-framework)
