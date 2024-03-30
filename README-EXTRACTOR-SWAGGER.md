# Extract Swagger Mapping

## Input parameters

### Path Parameter(s)

| Program Spec Table | Swagger.yaml               |
|--------------------|----------------------------|
| Parameters Name    | parameters.name            |
| ~~Value~~          | ~~parameters.schema.type~~ |
| Description        | /                          |
| Mandatory          | parameters.required        |

### Query Parameter(s)

N/A

### Request Body

| Program Spec Table | Swagger.yaml                                                                        |
|--------------------|-------------------------------------------------------------------------------------|
| Parameters Name    | requestBody...$ref(T) ->  components.schemas.T.properties[array].key                |
| Possible Values    | /                                                                                   |
| Description        | /                                                                                   |
| Mandatory          | requestBody.required AND ( requestBody...$ref(T) -> components.schemas.T.required ) |

## Business Rules & Logic (Description / Pseudo Code)

### Data Access Layer/Operation: Read

- GET/POST/PUT API
- Table rows follow-up Swagger(response data field)

| Program Spec Table      | Swagger.yaml                                                                                                                                                                          | Mybatis.xml   |
|-------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------------|
| Source Table            | /                                                                                                                                                                                     | db table name |
| Source Field Name       | /                                                                                                                                                                                     | db columns    |	
| Destination Entity Name | Constants -> Leave a mark "Response.payload#TODO" // (Insufficient information to implement) -> "Response.payload"/"Response.payload.documents[]"/"Response.payload. conversations[]" | /             |
| Destination Field Name  | requestBody...$ref(T) ->  components.schemas.T.properties[array].key                                                                                                                  | /             |
| Conversion Logic        | /                                                                                                                                                                                     | /             |

1. gen table rows by Swagger mapping [Destination Field Name]
2. extract mybatis data fill [Source Field Name], one to one mapping the Destination Field Name
3. fill [Source Table] by mybatis

### Data Access Layer/Operation: Write

- POST/PUT API
- Table rows follow-up Mybatis(db-schema columns)

| Program Spec Table     | Swagger.yaml                                                         | Mybatis.xml   |
|------------------------|----------------------------------------------------------------------|---------------|
| Source                 | Constants -> "Request Body"                                          | /             |
| Source Field Name      | requestBody...$ref(T) ->  components.schemas.T.properties[array].key | /             |	
| Destination            | /                                                                    | db table name |
| Destination Field Name | /                                                                    | db columns    |
| Conversion Logic       | /                                                                    | /             |

1. gen table rows by Mybatis mapping [Destination Field Name]
2. fill [Destination] by mybatis
3. extract swagger data fill [Source Field Name], one to one mapping the Destination Field Name

## Swagger API Data Compose

- loop paths prop locate $ref and loop components find mapping item
    - if components also locate $ref, continue loop itself and find mapping item