---
sidebar_position: 1
title: Introduction
slug: /
---

## What is Lokka?

Lokka lets you use Claude Desktop, or any MCP Client, to use natural language to accomplish things in your Microsoft 365 tenant through the Microsoft Graph API.

### Sample queries

Here are some examples of queries you can use with Lokka.

- `Create a new security group called 'Sales and HR' with a dynamic rule based on the department attribute.`
- `Find all the conditional access policies that haven't excluded the emergency access account`
- `Show me all the Intune device configuration policies assigned to the 'Call center' group`

You can ask Lokka to do anything that the Microsoft Graph API can do which includes support for Entra, Intune, Teams and SharePoint. The agent will only be able to perform the actions based on the permissions you grant it.

## What is MCP?

[Model Context Protocol](https://modelcontextprotocol.io/introduction) (MCP) is an open protocol that enables AI models to securely interact with local and remote resources through standardized server implementations.

Lokka is an implementation of the MCP protocol for the Microsoft Graph API.

![How does Lokka work?](./assets/how-does-lokka-mcp-server-work.png)

## Can I use this in production?

We recommend using Lokka in a test environment for exploration and testing purposes. The aim of this project is to provide a playground to expirement with using LLMs for Microsoft 365 administration tasks.

:::note

Lokka is not a production-ready solution and should not be used in a production environment. It is a proof of concept to demonstrate the capabilities of using LLMs for Microsoft 365 administration tasks.

:::

