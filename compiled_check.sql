with customers as (
    select * from analytics.staging.stg_customers
),

orders as (
    select * from analytics.staging.stg_orders
),

payments as (
    select * from analytics.staging.stg_payments
),

order_totals as (
    select
        order_id,
        sum(payment_amount) as total_paid,
        max(payment_date) as last_payment_date
    from payments
    group by order_id
),

customer_order_summary as (
    select
        o.customer_id,
        count(o.order_id) as total_orders,
        sum(t.total_paid) as lifetime_spend,
        max(t.last_payment_date) as last_order_date
    from orders o
    left join order_totals t on o.order_id = t.order_id
    group by o.customer_id
),

final as (
    select
        c.customer_id,
        c.customer_name,
        c.customer_email,
        coalesce(s.total_orders, 0) as total_orders,
        coalesce(s.lifetime_spend, 0) as lifetime_spend,
        s.last_order_date,
        case 
            when s.lifetime_spend >= 1000 then 'vip'
            when s.lifetime_spend > 0 then 'active'
            else 'churned'
        end as customer_segment
    from customers c
    left join customer_order_summary s on c.customer_id = s.customer_id
)

select * from final