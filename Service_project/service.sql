SELECT 
SUM( CASE WHEN service_id IS NULL THEN 1 ELSE 0 END ) AS SURVICE_ID_NULL,
SUM( CASE WHEN service_type_id IS NULL THEN 1 ELSE 0 END ) AS SURVICE_TYPE_ID_NULL,
SUM( CASE WHEN client_name IS NULL THEN 1 ELSE 0 END ) AS CLIENT_NAME_NULL
FROM service_data;

----------------------------
SELECT *
FROM service_data;

----------------------------

SELECT
 service_id,
 client_name,
 total_revenue,
(hours*hourly_rate) AS TOTAL_REV
FROM service_data
WHERE total_revenue <> (hours*hourly_rate);

--------------------------- 

CREATE VIEW V_FACT_SERVICES_ANALYSIS AS
SELECT
  S.SERVICE_ID,
  S.SERVICE_TYPE_ID,
  S.CLIENT_NAME,
  S.HOURS,
  CONVERT(DATE,S.SERVICE_DATE,103) AS SERVICE_DATE,
  S.SERVICE_TIME,
  S.HOURLY_RATE,
  S.TOTAL_REVENUE,
  S.DEPARTMENT,
  S.SERVICE_DESCRIPTION,
  B.COUNTRY,
  B.REGION
 FROM service_data S
 LEFT JOIN Branchw_data B
 ON S.branch_id = B.Branch_ID
 ;

 SELECT * FROM V_FACT_SERVICES_ANALYSIS;